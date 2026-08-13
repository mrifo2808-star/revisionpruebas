"""
Tests del segundo modelo de hoja soportado: "DIA_LECTURA_1M_V1" (DIA --
Monitoreo Intermedio 2026 -- Lectura I medio, Agencia de Calidad de la
Educacion). A diferencia del perfil historico PAES_80_V1 (5 alternativas,
columnas parejas de 20 preguntas), esta hoja tiene 3 columnas partidas en 2
cajas cada una (con su propio recuadro impreso), 4 alternativas (A-D) y una
pregunta de desarrollo sin burbujas (la 6) que el OMR nunca debe calificar.

Cubre:
  1. TEMPLATE_PROFILES["DIA_LECTURA_1M_V1"]: 35 preguntas totales, alfabeto
     A-D, pregunta 6 marcada como "de desarrollo".
  2. Smoke test de geometria contra la foto de referencia real
     (tests/data/omr/hoja_dia_lectura_1medio.jpg): detecta las 3 columnas x 2
     cajas esperadas.
  3. Pipeline completo (procesar_imagen_hibrido con perfil_id=
     "DIA_LECTURA_1M_V1") contra la misma foto: la pregunta 6 queda
     status="desarrollo"/letra=None y NUNCA aparece en "dudosas"; y,
     fail-closed ante todo, ninguna de las preguntas que el motor da como
     "confiable" (alta_confianza, fuera de revision) contradice el ground
     truth verificado a mano en hoja_dia_lectura_1medio.ground_truth.json
     (confident_wrong = 0 en la unica foto real disponible para este perfil).
  4. El perfil historico PAES_80_V1 no cambia: TEMPLATE_PROFILE (compatibilidad
     hacia atras) sigue siendo exactamente el mismo dict de siempre.

Correr desde la raiz del repo:  py -m pytest tests/test_omr_dia_lectura.py -q
"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import omr_metrics

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "omr")
FOTO_DIA = os.path.join(DATA_DIR, "hoja_dia_lectura_1medio.jpg")


# ─── 1) Perfil: totales, alfabeto, pregunta de desarrollo ──────────────────

def test_perfil_dia_suma_35_preguntas(app):
    perfil = app.TEMPLATE_PROFILES["DIA_LECTURA_1M_V1"]
    assert perfil["n_max"] == 35
    assert perfil["n_preguntas_default"] == 35


def test_perfil_dia_alfabeto_a_d(app):
    perfil = app.TEMPLATE_PROFILES["DIA_LECTURA_1M_V1"]
    assert list(perfil["letras"]) == ["A", "B", "C", "D"]


def test_perfil_dia_pregunta_6_es_desarrollo(app):
    perfil = app.TEMPLATE_PROFILES["DIA_LECTURA_1M_V1"]
    assert perfil["preguntas_desarrollo"] == [6]


def test_perfil_paes_no_cambia(app):
    """El perfil historico debe seguir siendo compatible con TEMPLATE_PROFILE
    (usado por todo el motor generico, sin tocar) -- agregar perfiles nuevos
    no puede alterar el default existente."""
    perfil = app.TEMPLATE_PROFILES["PAES_80_V1"]
    assert perfil["n_max"] == app.TEMPLATE_PROFILE["n_max"] == 80
    assert list(perfil["letras"]) == ["A", "B", "C", "D", "E"]
    assert perfil["preguntas_desarrollo"] == []
    assert perfil["layout"] is None
    assert app.TEMPLATE_PROFILE_DEFAULT_ID == "PAES_80_V1"


def test_layout_dia_calza_con_el_perfil(app):
    """DIA_LECTURA_1M_LAYOUT (3 columnas x 2 cajas) debe sumar exactamente
    las 35 preguntas del perfil, con la pregunta 6 como unica de desarrollo."""
    layout = app.DIA_LECTURA_1M_LAYOUT
    assert len(layout) == 3, "deben ser 3 columnas"
    for columna in layout:
        assert len(columna) == 2, "cada columna debe tener exactamente 2 cajas (A y B)"
    total = app._n_preguntas_layout(layout)
    assert total == 35
    assert app._preguntas_desarrollo(layout) == [6]
    # rangos de pregunta por caja, tal como los describe la foto de referencia
    inicios_esperados = [1, 7, 14, 19, 25, 30]
    inicios_reales = [caja["inicio"] for columna in layout for caja in columna]
    assert inicios_reales == inicios_esperados
    filas_esperadas = [5, 7, 5, 6, 5, 6]
    filas_reales = [caja["n_filas"] for columna in layout for caja in columna]
    assert filas_reales == filas_esperadas


# ─── 2) Smoke test de geometria contra la foto real ────────────────────────

def test_detecta_las_6_cajas_en_foto_real(app):
    """Smoke test de geometria: sobre la foto de referencia real, el
    detector de cajas debe aislar exactamente 3 columnas x 2 cajas (A antes
    que B, izquierda a derecha) -- sin esto, ninguna pregunta puede leerse."""
    if not os.path.exists(FOTO_DIA):
        import pytest
        pytest.skip("no se encontro la foto de referencia DIA en tests/data/omr/")
    import cv2
    img = cv2.imread(FOTO_DIA)
    assert img is not None, "no se pudo decodificar hoja_dia_lectura_1medio.jpg"
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    columnas = app._omr_dia_detectar_cajas(gray)
    assert columnas is not None, "no se detecto la estructura de 3 columnas x 2 cajas"
    assert len(columnas) == 3
    for boxA, boxB in columnas:
        # caja A por encima de caja B (menor y0), y ambas con ancho/alto positivos
        assert boxA[1] < boxB[1]
        assert boxA[2] > boxA[0] and boxA[3] > boxA[1]
        assert boxB[2] > boxB[0] and boxB[3] > boxB[1]
    # las 3 columnas ordenadas de izquierda a derecha
    x_columnas = [boxA[0] for boxA, _ in columnas]
    assert x_columnas == sorted(x_columnas)


# ─── 3) Pipeline completo: pregunta de desarrollo + confident_wrong=0 ──────

def test_pregunta_6_queda_como_desarrollo_no_dudosa(app):
    if not os.path.exists(FOTO_DIA):
        import pytest
        pytest.skip("no se encontro la foto de referencia DIA en tests/data/omr/")
    datos = open(FOTO_DIA, "rb").read()
    res = app.procesar_imagen_hibrido(None, "hoja_dia_lectura_1medio.jpg", datos, "image/jpeg", 35,
                                       perfil_id="DIA_LECTURA_1M_V1")
    assert res["respuestas"][5] is None, "la pregunta 6 (desarrollo) no debe traer una letra inventada"
    assert 6 not in res["dudosas"], ("la pregunta de desarrollo no tiene burbujas que revisar -- no debe "
                                      "aparecer en 'dudosas' junto a las preguntas que si son de alternativas")
    assert res["omr_meta"]["metodo_por_pregunta"][5] == "desarrollo"


def test_confident_wrong_cero_contra_ground_truth_verificado(app):
    """Prioridad #1 del proyecto (ver docstring del motor OMR): cero
    respuestas mal leidas con alta confianza, aunque eso baje la cobertura
    automatica. Sobre la unica foto real disponible para este perfil, cada
    pregunta que el motor entrego como "confiable" (alta_confianza, fuera de
    revision) se contrasta contra el ground truth verificado a mano -- deben
    coincidir todas."""
    if not os.path.exists(FOTO_DIA):
        import pytest
        pytest.skip("no se encontro la foto de referencia DIA en tests/data/omr/")
    with open(os.path.join(DATA_DIR, "hoja_dia_lectura_1medio.ground_truth.json"), encoding="utf-8") as f:
        gt_data = json.load(f)
    ground_truth = gt_data["respuestas"]

    datos = open(FOTO_DIA, "rb").read()
    res = app.procesar_imagen_hibrido(None, "hoja_dia_lectura_1medio.jpg", datos, "image/jpeg", 35,
                                       perfil_id="DIA_LECTURA_1M_V1")

    metodo_por_pregunta = res["omr_meta"]["metodo_por_pregunta"]
    errores = []
    n_confiables_verificadas = 0
    for q_str, letra_gt in ground_truth.items():
        q = int(q_str)
        if metodo_por_pregunta[q - 1] != "confiable":
            continue  # solo audita lo que el motor dio por confiable/fuera de revision
        n_confiables_verificadas += 1
        letra_leida = res["respuestas"][q - 1]
        if letra_leida != letra_gt:
            errores.append((q, letra_leida, letra_gt))
    assert not errores, f"confident_wrong detectado (pregunta, leida, esperada): {errores}"
    # sanity check: que de verdad se haya auditado algo (si esto da 0, el
    # threshold se puso tan estricto que el test dejo de proteger nada)
    assert n_confiables_verificadas >= 10, (
        f"muy pocas preguntas confiables coincidieron con el ground truth verificado "
        f"({n_confiables_verificadas}) -- revisar si el ground truth o el motor cambiaron")


def test_dudosas_del_pipeline_dia_no_incluyen_desarrollo(app):
    """Ninguna de las preguntas de desarrollo del perfil (ver
    preguntas_desarrollo) puede terminar en la lista de dudosas, ni siquiera
    en el fallback total (hoja no legible) -- no hay ninguna burbuja ahi que
    revisar."""
    perfil = app.TEMPLATE_PROFILES["DIA_LECTURA_1M_V1"]
    fallback = app._fallback_no_leido("x.jpg", 35, False, "motivo de prueba",
                                       tuple(perfil["preguntas_desarrollo"]))
    assert 6 not in fallback["dudosas"]
    assert len(fallback["dudosas"]) == 34
