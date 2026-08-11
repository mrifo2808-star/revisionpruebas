"""
Test de regresion permanente del motor OMR, corrido contra fotos reales (no
sinteticas) en tests/data/omr/. En particular fija en el tiempo el bug real
"USO EXCLUSIVO PARA ENSAYOS DE PRUEBAS": una foto real donde el bloque
RESPUESTAS detectado queda demasiado ancho y la 4ta banda (preguntas 61-80)
termina posicionada sobre el margen en blanco / texto vecino en vez de sobre
burbujas reales. Este test debe fallar si esa banda vuelve a devolver
alguna vez una respuesta "confiable" o "revisar_media" -- solo puede quedar
como revisar_geometria (sin letra, para revision manual).

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
    print("=== 2) hoja_uso_exclusivo.jpg: banda sobre 'USO EXCLUSIVO' nunca debe darse como confiable ===")
    with open(os.path.join(DATA_DIR, "hoja_uso_exclusivo.jpg"), "rb") as f:
        datos = f.read()
    res = app.procesar_imagen_hibrido(FakeCliente(), "hoja_uso_exclusivo.jpg", datos, "image/jpeg", 80,
                                       solo_respuestas=False)
    metodos = res["omr_meta"]["metodo_por_pregunta"]
    geo_conf = res["omr_meta"]["geometry_confidence_por_banda"]
    print("geometry_confidence_por_banda:", geo_conf)
    print("metodos preguntas 61-80:", set(metodos[60:80]))

    # La banda 4 (indice 3, preguntas 61-80) es la que en esta foto cae sobre
    # el margen en blanco / "USO EXCLUSIVO..." -- debe quedar marcada con baja
    # geometry_confidence y NINGUNA de sus preguntas puede darse como confiable.
    umbral = app.OMR_THRESHOLDS["MIN_GEOMETRY_CONFIDENCE"]
    assert geo_conf[3] < umbral, f"se esperaba baja geometry_confidence en la banda 4, dio {geo_conf[3]}"
    assert all(m == "revisar_geometria" for m in metodos[60:80]), \
        f"FALLO CRITICO: alguna pregunta 61-80 no quedo como revisar_geometria: {metodos[60:80]}"
    assert all(r is None for r in res["respuestas"][60:80]), \
        "FALLO CRITICO: se devolvio una letra para una pregunta sobre 'USO EXCLUSIVO' -- el bug reaparecio"

    # Las otras 3 bandas de ESTA MISMA foto (borrosa, con angulo real de celular)
    # deben seguir teniendo geometry_confidence aceptable -- el fix no debe
    # volverse tan estricto que desconfie de bandas legitimas solo por ruido/blur.
    for bi in (0, 1, 2):
        assert geo_conf[bi] >= umbral, f"banda {bi+1} no deberia quedar bajo el umbral (foto real valida): {geo_conf[bi]}"
    print("OK -- las 20 preguntas de la banda invadida por 'USO EXCLUSIVO' quedaron para revision manual\n")


if __name__ == "__main__":
    app = omr_metrics._cargar_app_module()
    test_calibracion_sin_regresion(app)
    test_uso_exclusivo_nunca_confiable(app)
    print("TODO PASO - regresion OMR con fotos reales OK")
