"""
Test de regresion permanente del motor OMR, corrido contra fotos reales (no
sinteticas) en tests/data/omr/. En particular fija en el tiempo el bug real
"USO EXCLUSIVO PARA ENSAYOS DE PRUEBAS": una foto real donde el bloque
RESPUESTAS detectado queda demasiado ancho y el contenido vecino ("USO
EXCLUSIVO...") competia como candidato de banda.

NOTA (2026-08-11, smoke test con fotos reales adicionales): la causa raiz de
este bug resulto ser mas profunda de lo que este test asumia originalmente.
No era que "la 4ta banda cae sobre el texto y por eso tiene baja confianza" --
era que la seleccion de bandas, al encontrar mas candidatos de los esperados,
se quedaba con LOS MAS ANCHOS, y el bloque de texto vecino podia ser mas
ancho que una columna real (angosta), asi que se descartaba una columna real
completa (la mas angosta) y el resto de las bandas quedaba UNA POSICION
CORRIDA -- una columna nunca se leia, sus respuestas venian en realidad de
la columna vecina bajo la etiqueta equivocada. Ese es el bug real (ver
tests/test_omr_band_selection.py para el test sintetico que lo fija).

Ya corregida la seleccion de bandas, en esta foto las 4 columnas quedan
correctamente identificadas (incluido el texto "USO EXCLUSIVO", que ahora
queda completamente fuera de las 4 bandas en vez de fusionado con alguna).
Este test ya no puede asumir que una banda especifica cae siempre sobre el
texto -- en su lugar fija la garantia real que importa: en esta foto dificil
(borrosa, angulo real de celular, geometria via UNIFORM_FALLBACK) NINGUNA
pregunta puede darse jamas como "confiable" (automatica sin revision).

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


def test_uso_exclusivo_nunca_confiable(app):
    print("=== 2) hoja_uso_exclusivo.jpg: ninguna pregunta puede darse como 'confiable' en esta foto dificil ===")
    with open(os.path.join(DATA_DIR, "hoja_uso_exclusivo.jpg"), "rb") as f:
        datos = f.read()
    res = app.procesar_imagen_hibrido(FakeCliente(), "hoja_uso_exclusivo.jpg", datos, "image/jpeg", 80,
                                       solo_respuestas=False)
    metodos = res["omr_meta"]["metodo_por_pregunta"]
    geo_conf = res["omr_meta"]["geometry_confidence_por_banda"]
    print("geometry_confidence_por_banda:", geo_conf)
    from collections import Counter
    print("metodos (conteo):", dict(Counter(metodos)))

    # Garantia real: en una foto dificil (borrosa, angulo real de celular),
    # cuya geometria completa cae en UNIFORM_FALLBACK, NINGUNA pregunta debe
    # darse jamas como "confiable" -- ni las que originalmente estaban sobre
    # texto, ni ninguna otra. Esto es lo que de verdad importa (una respuesta
    # "confiable" incorrecta), no en que banda especifica cae el texto.
    assert "confiable" not in metodos, \
        f"FALLO CRITICO: alguna pregunta se dio como 'confiable' en una foto con geometria via fallback: {Counter(metodos)}"

    # Deben existir exactamente 4 bandas -- ver tests/test_omr_band_selection.py
    # para el test dedicado de que las 4 columnas reales quedan correctamente
    # identificadas y el texto vecino excluido, en vez de fusionado con alguna
    # de ellas.
    #
    # NOTA (row lattice v4.1): antes de este cambio se exigia geometry_confidence
    # > 0 en las 4 bandas como proxy de "ninguna banda completamente vacia" --
    # ese numero se armaba solo con ancho/invariantes/fuente de columnas, sin
    # señal real de si las FILAS estaban bien alineadas. Ahora row_alignment_confidence
    # entra como quinto factor fail-closed (item 17 del plan): en esta foto real
    # cada banda tiene evidencia Hough genuina pero escasa (~11-24 circulos de
    # ~100 esperados) y los centros de columna vienen de UNIFORM_FALLBACK (no de
    # kmeans real), asi que el soporte fila-por-columna no alcanza el piso minimo
    # para confirmar alineacion -- geometry_confidence baja a 0.0 en las 4 bandas,
    # correctamente, en vez de fingir una confianza que la evidencia no respalda.
    # Lo que de verdad importa (ninguna respuesta "confiable" en una foto dificil)
    # ya se valido arriba y sigue intacto.
    assert len(geo_conf) == 4, f"se esperaban 4 bandas, se obtuvieron {len(geo_conf)}"
    assert all(0.0 <= g <= 1.0 for g in geo_conf), f"geometry_confidence fuera de rango: {geo_conf}"
    print("OK -- ninguna pregunta se dio como 'confiable' en esta foto dificil\n")


if __name__ == "__main__":
    app = omr_metrics._cargar_app_module()
    test_calibracion_sin_regresion(app)
    test_uso_exclusivo_nunca_confiable(app)
    print("TODO PASO - regresion OMR con fotos reales OK")
