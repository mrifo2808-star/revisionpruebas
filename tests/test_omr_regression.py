"""
Test de regresion permanente del motor OMR, corrido contra fotos reales (no
sinteticas) en tests/data/omr/. En particular fija en el tiempo el bug real
"USO EXCLUSIVO PARA ENSAYOS DE PRUEBAS": una foto real donde el bloque
RESPUESTAS detectado queda demasiado ancho y el contenido vecino ("USO
EXCLUSIVO...") competia como candidato de banda.

NOTA (2026-08-11, smoke test con fotos reales adicionales): la causa raiz de
este bug resulto ser mas profunda de lo que este test asumia originalmente,
y se revisito DOS VECES el mismo dia.

Primera revision: no era que "la 4ta banda cae sobre el texto y por eso
tiene baja confianza" -- era que la seleccion de bandas, al encontrar mas
candidatos de los esperados, se quedaba con LOS MAS ANCHOS, y el bloque de
texto vecino podia ser mas ancho que una columna real (angosta), asi que se
descartaba una columna real completa (la mas angosta) y el resto de las
bandas quedaba UNA POSICION CORRIDA. Ver tests/test_omr_band_selection.py
para el test sintetico de esa causa.

Segunda revision (misma fecha, tras un reporte de que la app en produccion
"no detecta ninguna respuesta" en una foto nueva): la causa raiz real era
TODAVIA mas arriba en el pipeline -- `omr_detectar_bloque_respuestas` podia
fusionar en un solo contorno la seccion de RESPUESTAS con secciones vecinas
(identificacion del alumno, N de folleto) via `cv2.morphologyEx(CLOSE,
kernel=9x9, iterations=2)`, cuyo alcance efectivo (~16-18px) puenteaba el
hueco real medido entre secciones DISTINTAS (13-14px) aunque fuera mayor que
el hueco DENTRO de una misma seccion (2-6px). El candidato resultante
(a veces casi la pagina completa) pasaba `_omr_candidato_tiene_header`
porque esa funcion solo confirmaba "hay ALGUNA barra oscura arriba" -- sin
distinguir la barra real "RESPUESTAS" de la barra, tambien real, de
"IDENTIFICACION DEL ESTUDIANTE" mas arriba. El resultado: la app leia
casillas de nombre/folleto como si fueran burbujas de respuesta. Corregido
con: iterations=1 en el cierre morfologico (dejaba de puentear el hueco
real entre secciones) + `_omr_header_seguido_de_grilla` (exige que la franja
inmediatamente debajo del header tenga la DENSIDAD de filas de una tabla de
burbujas real, no solo la presencia de una barra oscura).

Con esta segunda correccion, `hoja_uso_exclusivo.jpg` ya no cae a
UNIFORM_FALLBACK -- el bloque RESPUESTAS queda correctamente aislado con
evidencia Hough real en 3 de 4 bandas (antes ninguna pregunta podia darse
como "confiable" en esta foto). Se verifico que las 44 respuestas que ahora
se dan como confiables coinciden EXACTO, letra por letra, con el ground
truth ya validado de hoja_calibracion.png (que resulto ser la MISMA hoja de
respuestas -- incluso las preguntas en blanco/dudosas coinciden) -- ver
`test_uso_exclusivo_coincide_con_calibracion` mas abajo, que reemplaza la
garantia anterior ("nunca confiable") por una mas fuerte y mas real
(exactitud verificada, no solo ausencia de la palabra "confiable").

Correr desde la raiz del repo:  py tests/test_omr_regression.py
"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import omr_metrics

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "omr")


class FakeBlock:
    def __init__(self, text): self.type = "text"; self.text = text
class FakeMsg:
    def __init__(self, text): self.content = [FakeBlock(text)]; self.stop_reason = "end_turn"
class FakeMessages:
    def create(self, model, max_tokens, messages):
        return FakeMsg(json.dumps({"apellido_paterno": "", "apellido_materno": "",
                                    "nombres": "", "cedula": "", "nro_folleto": ""}))
class FakeCliente:
    def __init__(self): self.messages = FakeMessages()


def test_calibracion_sin_regresion(app):
    print("=== 1) hoja_calibracion.png: no debe haber regresion de exactitud ===")
    with open(os.path.join(DATA_DIR, "hoja_calibracion.ground_truth.json"), encoding="utf-8") as f:
        gt_data = json.load(f)
    gt = [gt_data["respuestas"].get(str(i)) for i in range(1, gt_data["n_preguntas"] + 1)]

    with open(os.path.join(DATA_DIR, "hoja_calibracion.png"), "rb") as f:
        datos = f.read()
    res = app.procesar_imagen_hibrido(FakeCliente(), "hoja_calibracion.png", datos, "image/png", 80,
                                       solo_respuestas=False)

    errores = [(i + 1, p, g) for i, (p, g) in enumerate(zip(res["respuestas"], gt)) if p is not None and p != g]
    print("aciertos directos:", sum(1 for p, g in zip(res["respuestas"], gt) if p == g and p is not None))
    print("errores en no-dudosas (debe ser []):", errores)
    assert not errores, f"FALLO: respuestas incorrectas dadas como no-dudosas: {errores}"
    assert res["omr_meta"]["n_geometry_error"] == 0, "la foto de calibracion no debe disparar geometry_error"
    print("OK\n")


def test_uso_exclusivo_coincide_con_calibracion(app):
    print("=== 2) hoja_uso_exclusivo.jpg: el bloque RESPUESTAS debe quedar aislado y las respuestas dadas "
          "como no-dudosas deben coincidir con el ground truth verificado de hoja_calibracion.png "
          "(misma hoja de respuestas) ===")
    with open(os.path.join(DATA_DIR, "hoja_calibracion.ground_truth.json"), encoding="utf-8") as f:
        gt_data = json.load(f)
    gt = [gt_data["respuestas"].get(str(i)) for i in range(1, gt_data["n_preguntas"] + 1)]

    with open(os.path.join(DATA_DIR, "hoja_uso_exclusivo.jpg"), "rb") as f:
        datos = f.read()
    res = app.procesar_imagen_hibrido(FakeCliente(), "hoja_uso_exclusivo.jpg", datos, "image/jpeg", 80,
                                       solo_respuestas=False)
    metodos = res["omr_meta"]["metodo_por_pregunta"]
    geo_conf = res["omr_meta"]["geometry_confidence_por_banda"]
    from collections import Counter
    print("geometry_confidence_por_banda:", geo_conf)
    print("metodos (conteo):", dict(Counter(metodos)))

    # Garantia real (nunca se relaja): ninguna respuesta que el motor da como
    # no-dudosa ("confiable") puede estar equivocada -- una "confiable"
    # incorrecta es la falla mas grave posible, mucho peor que mandar de mas
    # a revision manual.
    errores = [(i + 1, p, g) for i, (p, g) in enumerate(zip(res["respuestas"], gt))
               if p is not None and metodos[i] == "confiable" and p != g]
    print("errores en respuestas 'confiable' (debe ser []):", errores)
    assert not errores, f"FALLO CRITICO: respuestas incorrectas dadas como 'confiable': {errores}"

    # Con el bloque RESPUESTAS correctamente aislado (ver NOTA arriba: fix de
    # `omr_detectar_bloque_respuestas`/`_omr_header_seguido_de_grilla`), esta
    # foto dificil deja de caer entera en UNIFORM_FALLBACK -- al menos varias
    # bandas deben alcanzar evidencia real y dar respuestas confiables
    # correctas. Si esto vuelve a caer a 0, es señal de una regresion real en
    # la localizacion del bloque, no de que "la foto es dificil" (ya
    # demostramos que es legible).
    n_confiables = sum(1 for m in metodos if m == "confiable")
    assert n_confiables >= 30, \
        f"regresion: se esperaban varias decenas de respuestas confiables y correctas, se obtuvieron {n_confiables}"
    assert len(geo_conf) == 4, f"se esperaban 4 bandas, se obtuvieron {len(geo_conf)}"
    print(f"OK -- {n_confiables} respuestas confiables, todas correctas contra el ground truth\n")


if __name__ == "__main__":
    app = omr_metrics._cargar_app_module()
    test_calibracion_sin_regresion(app)
    test_uso_exclusivo_coincide_con_calibracion(app)
    print("TODO PASO - regresion OMR con fotos reales OK")
