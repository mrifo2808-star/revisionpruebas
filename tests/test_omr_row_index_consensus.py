"""
Tests del TRIPLE CHECK DE ÍNDICE DE FILA (v4.2, ver app_revisor.py, sección
"TRIPLE CHECK DE ÍNDICE DE FILA (v4.2, SHADOW MODE)"): refuerzo del error más
peligroso del OMR -- una respuesta visualmente bien leída pero asignada al
número de pregunta equivocado (ROW_INDEX_SHIFT). Agrega dos sensores
geométricos independientes al ROW LATTICE (sensor 1, ya existente) y un
combinador de consenso:

  SENSOR 2 -- omr_calcular_consenso_fase_global: consenso robusto (mediana
    leave-one-out, tolera hasta 1 banda corrida sobre 4) de la FASE/ORIGEN
    (y0) entre las bandas hermanas de la misma hoja impresa.
  SENSOR 3 -- omr_number_lattice_shift_hypothesis: en vez de estimar un
    origen propio para la franja de números (sensible a que falte justo el
    dígito que ancla el origen -- causa del falso positivo histórico del
    dígito "1", ver NUMBER_LATTICE_CROSSCHECK_HABILITADO), evalúa hipótesis
    de shift entero directamente contra el lattice de burbujas ya resuelto.
  CONSENSO -- omr_consenso_indice_fila: combina las tres señales sin dejar
    que ninguna por sí sola pueda remapear preguntas -- NUMBER es siempre
    auxiliar, nunca fuente primaria.

Todo esto opera en SHADOW MODE por defecto (ROW_INDEX_CONSENSUS_VETO_HABILITADO
= False): se calcula y expone siempre en diagnóstico, pero no cambia
geometry_confidence/geometry_state ni remapea una sola fila mientras el flag
siga apagado -- estos tests fijan ese comportamiento.

Sintéticos (sin fotos reales, control total sobre la verdad conocida), en la
misma línea que tests/test_omr_row_lattice.py y tests/test_omr_number_lattice.py.

Correr desde la raíz del repo:  py -m pytest tests/test_omr_row_index_consensus.py -q
"""
import os
import sys

import cv2
import numpy as np

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import omr_metrics

N_FILAS = 20
BH = 900.0
BW = 480.0
MARGEN_REAL = BH * 0.02
Y0_REAL = MARGEN_REAL
DY_REAL = (BH - 2 * MARGEN_REAL) / (N_FILAS - 1)
X_CENTERS_REAL = np.linspace(40, BW - 40, 5)


# ═══════════════════════ helpers (independientes, self-contained) ═════════════════════

def _generar_obs(rng, y0=Y0_REAL, dy=DY_REAL, x_centers=X_CENTERS_REAL,
                  jitter_y=0.6, jitter_x=0.6, drop_frac=0.0, drop_indices=(),
                  n_outliers=0, n_filas=N_FILAS):
    """Igual principio que tests/test_omr_row_lattice.py::_generar_obs, con la
    posibilidad adicional de tirar filas puntuales por índice (drop_indices),
    no solo una fracción aleatoria -- para los casos "fila intermedia/última
    ausente" que piden un índice concreto, no una probabilidad."""
    filas_y = [y0 + i * dy for i in range(n_filas)]
    obs = []
    for i, y_fila in enumerate(filas_y):
        if i in drop_indices or rng.random() < drop_frac:
            continue
        for xc in x_centers:
            obs.append((xc + rng.normal(0, jitter_x), y_fila + rng.normal(0, jitter_y)))
    for _ in range(n_outliers):
        obs.append((rng.uniform(0, BW), rng.uniform(0, BH)))
    return np.array(obs), np.array(filas_y)


def _row_errors(y_centers, filas_y_verdad, dy):
    return np.abs(np.asarray(y_centers) - np.asarray(filas_y_verdad)) / dy


N2_FILAS = 20
N2_BH = 900.0
N2_MARGEN = N2_BH * 0.02
N2_DY = (N2_BH - 2 * N2_MARGEN) / (N2_FILAS - 1)
BUBBLE_Y = N2_MARGEN + np.arange(N2_FILAS) * N2_DY


# ══════════════════════ SENSOR 1 (ROW LATTICE) -- casos complementarios ═══════════════
# Items 1,2,7,8,9 ya cubiertos por tests/test_omr_row_lattice.py (grilla
# perfecta, primera fila ausente, outliers, jitter, perspectiva residual).
# Acá se agregan los que faltan: fila intermedia/última ausente, 20%/30% de
# detecciones perdidas, y shift artificial +1/-1 (el mecanismo interno de
# autocorrección ROW_INDEX_SHIFT_corregido, ver omr_ajustar_lattice_vertical).

def test_caso3_fila_intermedia_ausente(app):
    print("=== caso 3: fila intermedia (indice 10) sin ninguna deteccion -- debe recuperarse por interpolacion ===")
    rng = np.random.default_rng(20)
    obs, filas_y = _generar_obs(rng, drop_indices=(10,))
    lat = app.omr_ajustar_lattice_vertical(obs, X_CENTERS_REAL, BH, N_FILAS, radio_banda=6.0)
    err = _row_errors(lat["y_centers"], filas_y, lat["dy"])
    print("row_error fila 11 (sin evidencia propia):", round(err[10], 4))
    assert err[10] < 0.35, f"FALLO: fila intermedia ausente quedo mal ubicada: {err[10]}"
    assert np.median(err) < 0.10
    print("OK\n")


def test_caso4_ultima_fila_ausente(app):
    print("=== caso 4: ultima fila (indice 19) sin ninguna deteccion -- no debe acortar/correr la grilla ===")
    rng = np.random.default_rng(21)
    obs, filas_y = _generar_obs(rng, drop_indices=(19,))
    lat = app.omr_ajustar_lattice_vertical(obs, X_CENTERS_REAL, BH, N_FILAS, radio_banda=6.0)
    err = _row_errors(lat["y_centers"], filas_y, lat["dy"])
    print("row_error ultima fila (sin evidencia propia):", round(err[-1], 4))
    assert err[-1] < 0.35, f"FALLO: ultima fila ausente quedo mal ubicada: {err[-1]}"
    print("OK\n")


def test_caso5_20pct_detecciones_perdidas(app):
    print("=== caso 5: 20% de las filas sin ninguna deteccion -- row_alignment_confidence debe seguir siendo util ===")
    rng = np.random.default_rng(22)
    obs, filas_y = _generar_obs(rng, drop_frac=0.20, jitter_y=0.6, jitter_x=0.6)
    lat = app.omr_ajustar_lattice_vertical(obs, X_CENTERS_REAL, BH, N_FILAS, radio_banda=6.0)
    err = _row_errors(lat["y_centers"], filas_y, lat["dy"])
    print("median row_error:", round(np.median(err), 4), "row_alignment_confidence:", round(lat["row_alignment_confidence"], 3))
    assert np.median(err) < 0.15
    assert lat["row_alignment_confidence"] > 0.3, "con 80% de evidencia real, la banda no deberia caer a confianza nula"
    print("OK\n")


def test_caso6_30pct_detecciones_perdidas(app):
    print("=== caso 6: 30% de las filas sin ninguna deteccion -- debe degradar, no colapsar de forma catastrofica ===")
    rng = np.random.default_rng(23)
    obs, filas_y = _generar_obs(rng, drop_frac=0.30, jitter_y=0.6, jitter_x=0.6)
    lat = app.omr_ajustar_lattice_vertical(obs, X_CENTERS_REAL, BH, N_FILAS, radio_banda=6.0)
    err = _row_errors(lat["y_centers"], filas_y, lat["dy"])
    print("median row_error:", round(np.median(err), 4), "max:", round(err.max(), 4),
          "row_alignment_confidence:", round(lat["row_alignment_confidence"], 3))
    assert np.median(err) < 0.20
    assert err.max() < 0.45, "ninguna fila puede cruzar la mitad de camino a la vecina aunque falte 30% de evidencia"
    print("OK\n")


def test_caso10_shift_artificial_mas_uno(app):
    print("=== caso 10: la grilla REAL esta corrida +1 fila respecto del margen esperado por plantilla -- "
          "debe autocorregirse (ROW_INDEX_SHIFT_corregido), no quedar anclada al margen equivocado ===")
    rng = np.random.default_rng(24)
    y0_corrido = Y0_REAL + DY_REAL  # el margen impreso real quedo una fila mas abajo de lo esperado
    obs, filas_y = _generar_obs(rng, y0=y0_corrido, jitter_y=0.5, jitter_x=0.5)
    lat = app.omr_ajustar_lattice_vertical(obs, X_CENTERS_REAL, BH, N_FILAS, radio_banda=6.0)
    err = _row_errors(lat["y_centers"], filas_y, lat["dy"])
    print("y0 recuperado:", round(lat["y0"], 2), " y0 real:", round(y0_corrido, 2),
          " violaciones:", lat["diagnostico"]["violaciones"])
    assert "ROW_INDEX_SHIFT_corregido" in lat["diagnostico"]["violaciones"], \
        "el mecanismo interno de autocorreccion no detecto el shift +1 sintetico"
    assert np.median(err) < 0.15, f"tras corregir, el error contra la grilla REAL (corrida) debe ser chico: {np.median(err)}"
    print("OK\n")


def test_caso11_shift_artificial_menos_uno(app):
    print("=== caso 11: la grilla REAL esta corrida -1 fila respecto del margen esperado por plantilla ===")
    rng = np.random.default_rng(25)
    y0_corrido = Y0_REAL - DY_REAL
    obs, filas_y = _generar_obs(rng, y0=y0_corrido, jitter_y=0.5, jitter_x=0.5)
    lat = app.omr_ajustar_lattice_vertical(obs, X_CENTERS_REAL, BH, N_FILAS, radio_banda=6.0)
    err = _row_errors(lat["y_centers"], filas_y, lat["dy"])
    print("y0 recuperado:", round(lat["y0"], 2), " y0 real:", round(y0_corrido, 2),
          " violaciones:", lat["diagnostico"]["violaciones"])
    assert "ROW_INDEX_SHIFT_corregido" in lat["diagnostico"]["violaciones"], \
        "el mecanismo interno de autocorreccion no detecto el shift -1 sintetico"
    assert np.median(err) < 0.15, f"tras corregir, el error contra la grilla REAL (corrida) debe ser chico: {np.median(err)}"
    print("OK\n")


# ═══════════════════════ SENSOR 2 -- GLOBAL ROW PHASE CONSENSUS ═══════════════════════

def test_caso12_cuatro_bandas_alineadas(app):
    print("=== caso 12: 4 bandas con y0 casi identico -- consenso debe confirmar 'sin shift' con confianza alta ===")
    y0s = [100.0, 100.5, 99.7, 100.2]
    ev = [True, True, True, True]
    res = app.omr_calcular_consenso_fase_global(y0s, N2_DY, ev)
    for r in res:
        assert r["row_shift_candidate"] == 0
        assert r["row_shift_confidence"] > 0.6
        assert r["evidencia_suficiente"] is True
    print("OK\n")


def test_caso13_banda1_shift_mas_uno(app):
    print("=== caso 13: banda 1 (indice 0) corrida +1 fila -- debe identificarse como outlier sin arrastrar a las otras ===")
    y0s = [100.0 + N2_DY, 100.0, 100.3, 99.8]
    ev = [True, True, True, True]
    res = app.omr_calcular_consenso_fase_global(y0s, N2_DY, ev)
    assert res[0]["row_shift_candidate"] == 1
    assert res[0]["row_shift_confidence"] > 0.6
    for i in (1, 2, 3):
        assert res[i]["row_shift_candidate"] == 0
        assert res[i]["row_shift_confidence"] > 0.6, f"la banda sana {i} no deberia perder confianza por el outlier"
    print("OK\n")


def test_caso14_banda2_shift_menos_uno(app):
    print("=== caso 14: banda 2 (indice 1) corrida -1 fila ===")
    y0s = [100.0, 100.0 - N2_DY, 100.2, 99.9]
    ev = [True, True, True, True]
    res = app.omr_calcular_consenso_fase_global(y0s, N2_DY, ev)
    assert res[1]["row_shift_candidate"] == -1
    assert res[1]["row_shift_confidence"] > 0.6
    for i in (0, 2, 3):
        assert res[i]["row_shift_candidate"] == 0
    print("OK\n")


def test_caso15_banda3_ruidosa(app):
    print("=== caso 15: banda 3 (indice 2) con un offset que NO es un shift entero limpio (ruido, no shift real) -- "
          "no debe confirmarse ningun shift con confianza, debe quedar 'dudosa' ===")
    y0s = [100.0, 100.2, 100.0 + N2_DY * 0.4, 99.9]
    ev = [True, True, True, True]
    res = app.omr_calcular_consenso_fase_global(y0s, N2_DY, ev)
    print(res[2])
    assert res[2]["row_shift_confidence"] < 0.3, "un offset no-entero (ruido) no deberia poder confirmar nada con fuerza"
    print("OK\n")


def test_caso16_banda4_sin_evidencia_suficiente(app):
    print("=== caso 16: banda 4 (indice 3) sin evidencia propia -- no puede ser jueza ni acusada ===")
    y0s = [100.0, 100.2, 99.8, 250.0]
    ev = [True, True, True, False]
    res = app.omr_calcular_consenso_fase_global(y0s, N2_DY, ev)
    assert res[3]["evidencia_suficiente"] is False
    assert res[3]["row_shift_candidate"] == 0
    for i in (0, 1, 2):
        assert res[i]["row_shift_candidate"] == 0
        assert res[i]["row_shift_confidence"] > 0.6, "la banda sin evidencia no deberia contaminar a las demas"
    print("OK\n")


def test_caso17_tres_de_acuerdo_una_outlier(app):
    print("=== caso 17: tres bandas de acuerdo + una outlier clara -- la mediana leave-one-out no debe dejarse "
          "arrastrar por el outlier al juzgar a las tres bandas sanas ===")
    y0s = [100.0, 100.3, 99.7, 100.0 + N2_DY]
    ev = [True, True, True, True]
    res = app.omr_calcular_consenso_fase_global(y0s, N2_DY, ev)
    assert res[3]["row_shift_candidate"] == 1 and res[3]["row_shift_confidence"] > 0.6
    for i in (0, 1, 2):
        assert res[i]["row_shift_candidate"] == 0 and res[i]["row_shift_confidence"] > 0.6
    print("OK\n")


# ═══════════════════════ SENSOR 3 -- NUMBER LATTICE por hipotesis de shift ════════════

def test_caso18_fila1_invisible(app):
    print("=== caso 18: fila 1 (indice 0) sin ninguna tinta detectada -- no debe inventar un shift ===")
    r = app.omr_number_lattice_shift_hypothesis(BUBBLE_Y[1:], BUBBLE_Y, N2_DY)
    print(r)
    assert r["shift_candidate"] == 0
    print("OK\n")


def test_caso19_numero_1_muy_debil(app):
    print("=== caso 19: el mismo escenario historico real -- digito '1' demasiado fino para detectarse ===")
    rng = np.random.default_rng(30)
    obs = BUBBLE_Y[1:] + rng.normal(0, 0.4, size=N2_FILAS - 1)
    r = app.omr_number_lattice_shift_hypothesis(obs, BUBBLE_Y, N2_DY)
    print(r)
    assert r["shift_candidate"] == 0, "FALLO CRITICO: el numero '1' debil produjo un shift falso (regresion historica)"
    print("OK\n")


def test_caso20_numero_11_muy_debil(app):
    print("=== caso 20: numero de fila intermedia (11, indice 10) debil/ausente ===")
    idx = list(range(N2_FILAS))
    idx.remove(10)
    r = app.omr_number_lattice_shift_hypothesis(BUBBLE_Y[idx], BUBBLE_Y, N2_DY)
    print(r)
    assert r["shift_candidate"] == 0
    print("OK\n")


def test_caso21_varias_observaciones_ausentes(app):
    print("=== caso 21: varias filas (5 de 20) sin observacion -- no debe afirmar shift ===")
    rng = np.random.default_rng(31)
    mask = np.ones(N2_FILAS, dtype=bool)
    mask[[2, 6, 9, 14, 18]] = False
    obs = BUBBLE_Y[mask] + rng.normal(0, 0.4, size=mask.sum())
    r = app.omr_number_lattice_shift_hypothesis(obs, BUBBLE_Y, N2_DY)
    print(r)
    assert r["shift_candidate"] == 0
    print("OK\n")


def test_caso22_30pct_numeros_ausentes(app):
    print("=== caso 22: 30% de los numeros ausentes -- sensor auxiliar, esta bien que quede sin evidencia clara ===")
    rng = np.random.default_rng(32)
    mask = rng.random(N2_FILAS) > 0.30
    obs = BUBBLE_Y[mask] + rng.normal(0, 0.4, size=mask.sum())
    r = app.omr_number_lattice_shift_hypothesis(obs, BUBBLE_Y, N2_DY)
    print(r)
    assert r["shift_candidate"] == 0, "con evidencia parcial y sin shift real, jamas debe inventar uno"
    print("OK\n")


def test_caso23_artefacto_antes_de_primera_fila(app):
    print("=== caso 23: un artefacto de tinta antes de la fila 1 (fuera de la tabla) no debe torcer el veredicto ===")
    rng = np.random.default_rng(33)
    obs = np.concatenate([BUBBLE_Y + rng.normal(0, 0.3, size=N2_FILAS), [BUBBLE_Y[0] - 2.3 * N2_DY]])
    r = app.omr_number_lattice_shift_hypothesis(obs, BUBBLE_Y, N2_DY)
    print(r)
    assert r["shift_candidate"] == 0
    print("OK\n")


def test_caso24_artefacto_entre_filas(app):
    print("=== caso 24: un artefacto de tinta a mitad de camino entre dos filas reales no debe torcer el veredicto ===")
    rng = np.random.default_rng(34)
    obs = np.concatenate([BUBBLE_Y + rng.normal(0, 0.3, size=N2_FILAS), [BUBBLE_Y[7] + N2_DY / 2]])
    r = app.omr_number_lattice_shift_hypothesis(obs, BUBBLE_Y, N2_DY)
    print(r)
    assert r["shift_candidate"] == 0
    print("OK\n")


def test_caso25_shift_verdadero_mas_uno(app):
    print("=== caso 25: numeros impresos realmente una fila mas abajo que el lattice de burbujas (20/20, limpio) ===")
    r = app.omr_number_lattice_shift_hypothesis(BUBBLE_Y + N2_DY, BUBBLE_Y, N2_DY)
    print(r)
    assert r["shift_candidate"] == 1
    assert r["shift_confidence"] >= 0.6, "con evidencia completa y limpia, la confianza deberia ser alta"
    print("OK\n")


def test_caso26_shift_verdadero_menos_uno(app):
    print("=== caso 26: numeros impresos realmente una fila mas arriba que el lattice de burbujas (20/20, limpio) ===")
    r = app.omr_number_lattice_shift_hypothesis(BUBBLE_Y - N2_DY, BUBBLE_Y, N2_DY)
    print(r)
    assert r["shift_candidate"] == -1
    assert r["shift_confidence"] >= 0.6
    print("OK\n")


def test_caso27_number_equivocado_con_row_correcto(app):
    print("=== caso 27: franja de numeros corrupta/no relacionada con la grilla real (ruido) -- "
          "no debe afirmar ningun shift con confianza, aunque el ROW LATTICE (no evaluado aca) este perfecto ===")
    rng = np.random.default_rng(35)
    obs_ruido = rng.uniform(0, N2_BH, size=12)
    r = app.omr_number_lattice_shift_hypothesis(obs_ruido, BUBBLE_Y, N2_DY)
    print(r)
    assert r["shift_candidate"] == 0
    assert r["shift_confidence"] < 0.3
    print("OK\n")


def test_ruido_puro_nunca_produce_shift_confiado(app):
    print("=== estres: 200 semillas de ruido puro -- jamas debe emerger un shift_candidate != 0 ===")
    falsos = 0
    for seed in range(200):
        rng = np.random.default_rng(4000 + seed)
        n = int(rng.integers(4, 15))
        obs = rng.uniform(0, N2_BH, size=n)
        r = app.omr_number_lattice_shift_hypothesis(obs, BUBBLE_Y, N2_DY)
        if r["shift_candidate"] != 0:
            falsos += 1
    print("falsos positivos en 200 pruebas de ruido puro:", falsos)
    assert falsos == 0, f"FALLO CRITICO: el sensor NUMBER afirmo un shift sobre ruido puro en {falsos}/200 pruebas"
    print("OK\n")


# ═══════════════════════════════ CONSENSO ══════════════════════════════════════════════

def _global_diag(shift=0, conf=0.9, err=0.02, suf=True):
    return {"row_shift_candidate": shift, "row_shift_confidence": conf, "global_phase_error": err,
            "evidencia_suficiente": suf}


def _number_diag(shift=0, conf=0.9, err=0.05, suf=True):
    return {"shift_candidate": shift, "shift_confidence": conf, "number_phase_error": err,
            "evidencia_suficiente": suf}


def test_caso28_number_incorrecto_no_genera_falso_shift(app):
    print("=== caso 28 (CRITICO): ROW + GLOBAL correctos, NUMBER fuerte pero EQUIVOCADO -- "
          "jamas puede generarse un shift falso ===")
    c = app.omr_consenso_indice_fila("clean", _global_diag(0, 0.9), _number_diag(1, 0.9))
    print(c)
    assert c["possible_row_index_shift"] is False, "FALLO CRITICO: NUMBER equivocado genero un shift falso"
    assert c["recomendacion"] == "MANTENER"
    assert c["row_shift_candidate"] == 0
    print("OK\n")


def test_caso29_row_number_correctos_global_dudoso(app):
    print("=== caso 29: ROW + NUMBER correctos, GLOBAL dudoso (poca confianza, sin shift claro) -- se mantiene ===")
    c = app.omr_consenso_indice_fila("clean", _global_diag(0, 0.4, suf=True), _number_diag(0, 0.8))
    print(c)
    assert c["possible_row_index_shift"] is False
    assert c["recomendacion"] == "MANTENER"
    print("OK\n")


def test_caso30_global_number_fuertes_row_dudoso(app):
    print("=== caso 30: GLOBAL + NUMBER fuertes de acuerdo en shift +1, ROW dudoso (ambiguo) -- shift probable, revisar ===")
    c = app.omr_consenso_indice_fila("ambiguous", _global_diag(1, 0.8), _number_diag(1, 0.75))
    print(c)
    assert c["possible_row_index_shift"] is True
    assert c["recomendacion"] == "SHIFT_PROBABLE_REVISAR"
    assert c["row_shift_candidate"] == 1
    assert c["sensor_agreement"] is True
    print("OK\n")


def test_caso31_los_tres_dudosos(app):
    print("=== caso 31: ROW ambiguo, GLOBAL debil, NUMBER debil -- FAIL CLOSED, enviar a revision ===")
    c = app.omr_consenso_indice_fila("ambiguous", _global_diag(1, 0.4), _number_diag(1, 0.4))
    print(c)
    assert c["possible_row_index_shift"] is True
    assert c["recomendacion"] == "FAIL_CLOSED_REVISAR"
    print("OK\n")


def test_caso32_los_tres_de_acuerdo(app):
    print("=== caso 32: ROW ya se autocorrigio + GLOBAL y NUMBER fuertes de acuerdo -- shift altamente probable, "
          "pero el resultado sigue siendo 'revisar', nunca un auto-shift ===")
    c = app.omr_consenso_indice_fila("corrected", _global_diag(1, 0.85), _number_diag(1, 0.8))
    print(c)
    assert c["possible_row_index_shift"] is True
    assert c["recomendacion"] == "SHIFT_PROBABLE_REVISAR"
    assert c["row_shift_candidate"] == 1
    print("OK\n")


def test_caso33_dos_fuertes_vs_uno_debil_contradictorio(app):
    print("=== caso 33: GLOBAL y NUMBER fuertes de acuerdo en 'sin shift', pero ROW mismo quedo ambiguo -- "
          "el riesgo propio de ROW no se descarta solo porque los otros dos digan 0 (mas conservador que A/D "
          "a proposito: ROW es el sensor primario) -- debe fallar cerrado, nunca inventar 'todo bien' ni un shift ===")
    c = app.omr_consenso_indice_fila("ambiguous", _global_diag(0, 0.9), _number_diag(0, 0.9))
    print(c)
    assert c["row_shift_candidate"] == 0, "no debe inventar un shift concreto"
    assert c["recomendacion"] == "FAIL_CLOSED_REVISAR"
    print("OK\n")


def test_caso_a_ignora_number_debil_discrepante(app):
    print("=== caso A: ROW limpio + GLOBAL fuerte en 0 + NUMBER debil discrepante -- se ignora NUMBER, se mantiene ===")
    c = app.omr_consenso_indice_fila("clean", _global_diag(0, 0.9), _number_diag(1, 0.4))
    print(c)
    assert c["possible_row_index_shift"] is False
    assert c["recomendacion"] == "MANTENER"
    print("OK\n")


def test_caso_d_number_sin_evidencia_no_penaliza(app):
    print("=== caso D: ROW limpio + GLOBAL fuerte en 0 + NUMBER sin evidencia -- la ausencia no penaliza ===")
    c = app.omr_consenso_indice_fila("clean", _global_diag(0, 0.9), _number_diag(0, 0.0, suf=False))
    print(c)
    assert c["possible_row_index_shift"] is False
    assert c["recomendacion"] == "MANTENER"
    print("OK\n")


def test_sin_ninguna_senal_se_mantiene(app):
    print("=== default: sin evidencia de ningun sensor (foto dificil, GLOBAL y NUMBER sin evidencia suficiente, "
          "ROW limpio pero sin confianza alta) -- no debe marcarse nada para revision solo por falta de evidencia ===")
    c = app.omr_consenso_indice_fila("clean", _global_diag(0, 0.0, suf=False), _number_diag(0, 0.0, suf=False))
    print(c)
    assert c["possible_row_index_shift"] is False
    assert c["recomendacion"] == "MANTENER"
    print("OK\n")


def test_conflicto_real_entre_global_y_number(app):
    print("=== dos sensores fuertes en desacuerdo entre si (GLOBAL dice +1, NUMBER dice -1) -- fail closed, "
          "ninguno tiene autoridad para desempatar al otro ===")
    c = app.omr_consenso_indice_fila("clean", _global_diag(1, 0.8), _number_diag(-1, 0.8))
    print(c)
    assert c["sensor_conflict"] is True
    assert c["possible_row_index_shift"] is True
    assert c["recomendacion"] == "FAIL_CLOSED_REVISAR"
    print("OK\n")


# ═══════════════════ SHADOW MODE -- verificacion de que el flag por defecto no afecta nada ═══════════════

def test_flag_veto_off_por_defecto(app):
    print("=== ROW_INDEX_CONSENSUS_VETO_HABILITADO debe estar OFF por defecto (shadow mode) ===")
    assert app.ROW_INDEX_CONSENSUS_VETO_HABILITADO is False, \
        "el veto del triple-check se habilito sin la evidencia real que pide el plan -- debe quedar en shadow mode"
    print("OK\n")


def _banda_circulos(img, x0, n_filas=20, n_cols=5, radio=8, spacing_x=28, spacing_y=None, y0=None):
    spacing_y = spacing_y if spacing_y is not None else DY_REAL
    y0 = y0 if y0 is not None else Y0_REAL
    for row in range(n_filas):
        for col in range(n_cols):
            cx = int(round(x0 + 20 + col * spacing_x))
            cy = int(round(y0 + row * spacing_y))
            if 0 <= cy < img.shape[0]:
                cv2.circle(img, (cx, cy), radio, 90, 2)


def test_integracion_shadow_mode_no_cambia_geometry_confidence(app):
    print("=== integracion: omr_ajustar_grilla con 4 bandas sinteticas (una con su grilla corrida +1 fila) -- "
          "con el flag en shadow mode, geometry_confidence NO debe verse afectado por el triple-check, "
          "y el diagnostico row_index_consensus_por_banda debe quedar expuesto igual ===")
    n_filas = 20
    band_w = 170
    bh = int(Y0_REAL + n_filas * DY_REAL + 40)
    bw = band_w * 4 + 20
    img = np.full((bh, bw), 255, dtype=np.uint8)
    bands = []
    for bi in range(4):
        x0 = 10 + bi * band_w
        x1 = x0 + band_w - 10
        bands.append((x0, x1))
        y0_banda = Y0_REAL + (DY_REAL if bi == 2 else 0.0)  # banda 2 (indice 2) corrida +1 fila
        _banda_circulos(img, x0, n_filas=n_filas, y0=y0_banda)

    salida = app.omr_ajustar_grilla(img, bands, n_filas=n_filas, bandas_fabricadas=False)
    assert len(salida) == 9, f"omr_ajustar_grilla debe devolver 9 elementos (se agrego row_index_consensus_por_banda): {len(salida)}"
    (y_centers, x_centers, radio, geometry_confidence, geometry_violaciones, geometry_source,
     row_align_conf, row_source, row_index_consensus) = salida
    print("geometry_confidence:", [round(g, 3) for g in geometry_confidence])
    print("row_index_consensus banda 2:", row_index_consensus[2])
    assert len(row_index_consensus) == 4
    for c in row_index_consensus:
        for k in ("row_shift_candidate", "row_shift_confidence", "global_phase_error", "number_phase_error",
                   "sensor_agreement", "sensor_conflict", "band_outlier", "possible_row_index_shift", "recomendacion"):
            assert k in c, f"falta la clave {k} en el diagnostico de consenso"

    # SHADOW MODE: aunque el consenso detecte algo en la banda 2, ninguna
    # violacion de consenso sin el sufijo "_ignorado_flag_off" puede aparecer
    # mientras el flag este apagado -- y geometry_confidence de bandas con
    # geometria sana no puede haber sido forzado a 0 por el triple-check.
    for bi, violaciones in enumerate(geometry_violaciones):
        for v in violaciones:
            if v.startswith("ROW_INDEX_CONSENSUS_"):
                assert v.endswith("_ignorado_flag_off"), \
                    f"banda {bi}: el veto del triple-check actuo con el flag apagado: {v}"
    for bi in range(4):
        assert geometry_confidence[bi] > 0.5, \
            f"banda {bi}: geometry_confidence no deberia degradarse por el triple-check en shadow mode: {geometry_confidence[bi]}"
    print("OK\n")


if __name__ == "__main__":
    app = omr_metrics._cargar_app_module()
    test_caso3_fila_intermedia_ausente(app)
    test_caso4_ultima_fila_ausente(app)
    test_caso5_20pct_detecciones_perdidas(app)
    test_caso6_30pct_detecciones_perdidas(app)
    test_caso10_shift_artificial_mas_uno(app)
    test_caso11_shift_artificial_menos_uno(app)
    test_caso12_cuatro_bandas_alineadas(app)
    test_caso13_banda1_shift_mas_uno(app)
    test_caso14_banda2_shift_menos_uno(app)
    test_caso15_banda3_ruidosa(app)
    test_caso16_banda4_sin_evidencia_suficiente(app)
    test_caso17_tres_de_acuerdo_una_outlier(app)
    test_caso18_fila1_invisible(app)
    test_caso19_numero_1_muy_debil(app)
    test_caso20_numero_11_muy_debil(app)
    test_caso21_varias_observaciones_ausentes(app)
    test_caso22_30pct_numeros_ausentes(app)
    test_caso23_artefacto_antes_de_primera_fila(app)
    test_caso24_artefacto_entre_filas(app)
    test_caso25_shift_verdadero_mas_uno(app)
    test_caso26_shift_verdadero_menos_uno(app)
    test_caso27_number_equivocado_con_row_correcto(app)
    test_ruido_puro_nunca_produce_shift_confiado(app)
    test_caso28_number_incorrecto_no_genera_falso_shift(app)
    test_caso29_row_number_correctos_global_dudoso(app)
    test_caso30_global_number_fuertes_row_dudoso(app)
    test_caso31_los_tres_dudosos(app)
    test_caso32_los_tres_de_acuerdo(app)
    test_caso33_dos_fuertes_vs_uno_debil_contradictorio(app)
    test_caso_a_ignora_number_debil_discrepante(app)
    test_caso_d_number_sin_evidencia_no_penaliza(app)
    test_sin_ninguna_senal_se_mantiene(app)
    test_conflicto_real_entre_global_y_number(app)
    test_flag_veto_off_por_defecto(app)
    test_integracion_shadow_mode_no_cambia_geometry_confidence(app)
    print("TODO PASO - triple check de indice de fila (v4.2) OK")
