"""
Tests del NUMBER LATTICE (v4.1, ver omr_ajustar_number_lattice/
omr_crosscheck_bubble_number en app_revisor.py): segundo sensor geometrico
INDEPENDIENTE de las burbujas, basado en la franja de números de fila
impresos a la izquierda de la columna A -- sin OCR, solo presencia de tinta
(mismo principio que _omr_encontrar_barra_encabezado).

Sintéticos (sin fotos reales): construimos directamente la franja en escala
de grises con "tinta" simulada en posiciones Y conocidas, para tener control
total sobre la verdad contra la que medir.

IMPORTANTE (ver NUMBER_LATTICE_CROSSCHECK_HABILITADO en app_revisor.py): el
veto del crosscheck sobre geometry_confidence está OFF por defecto -- se
demostró un falso positivo en la foto de calibración real (el dígito "1" de
la fila 1 es demasiado delgado para producir tinta detectable a la
resolución de detección, corriendo el lattice de números una fila completa
contra el de burbujas, que sí era correcto). Estos tests fijan ese
comportamiento: el sensor existe y se calcula, pero no debe poder degradar
geometry_confidence mientras el flag siga apagado.

Correr desde la raíz del repo:  py -m pytest tests/test_omr_number_lattice.py -q
"""
import os
import sys

import numpy as np

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import omr_metrics

N_FILAS = 20
BH = 900
BW = 200
MARGEN = BH * 0.02
DY = (BH - 2 * MARGEN) / (N_FILAS - 1)


def _franja_sintetica(rng, y0=MARGEN, dy=DY, drop_frac=0.0, ancho_ink=3, jitter=0.4, franja_w=8):
    """Construye una franja en escala de grises (uint8) con "tinta" (oscura)
    en cada fila esperada, salvo las que se hagan caer en drop_frac -- simula
    números de fila reales impresos, sin dibujar dígitos de verdad (no hace
    falta: el detector solo mide presencia de tinta, nunca OCR)."""
    franja = np.full((BH, franja_w), 230, dtype=np.uint8)
    filas_y = []
    for i in range(N_FILAS):
        y = y0 + i * dy
        filas_y.append(y)
        if rng.random() < drop_frac:
            continue
        yc = int(round(y + rng.normal(0, jitter)))
        y_ini, y_fin = max(0, yc - ancho_ink // 2), min(BH, yc + ancho_ink // 2 + 1)
        franja[y_ini:y_fin, :] = 40
    return franja, np.array(filas_y)


def test_number_lattice_recupera_periodo_y_origen(app):
    print("=== franja sintetica limpia: number lattice debe recuperar y0/dy con precision ===")
    rng = np.random.default_rng(10)
    franja, filas_y = _franja_sintetica(rng)
    obs = app._omr_detectar_number_y_obs(franja)
    print("n_obs:", len(obs))
    y_sorted = np.sort(obs)
    dy_local, n_clusters, _diag = app._omr_estimar_dy_robusto(y_sorted, BH, N_FILAS, BH / N_FILAS)
    print("dy_local:", dy_local, "esperado:", DY)
    assert abs(dy_local - DY) / DY < 0.05
    assert n_clusters >= N_FILAS - 2
    print("OK\n")


def test_number_lattice_tolera_filas_faltantes(app):
    print("=== franja con 25% de numeros faltantes (tinta debil/borrosa): dy sigue siendo recuperable ===")
    rng = np.random.default_rng(11)
    franja, filas_y = _franja_sintetica(rng, drop_frac=0.25)
    obs = app._omr_detectar_number_y_obs(franja)
    y_sorted = np.sort(obs)
    dy_local, n_clusters, _diag = app._omr_estimar_dy_robusto(y_sorted, BH, N_FILAS, BH / N_FILAS)
    print("n_obs:", len(obs), "dy_local:", dy_local, "n_clusters:", n_clusters)
    assert abs(dy_local - DY) / DY < 0.08
    print("OK\n")


def test_number_lattice_ignora_artefacto_de_borde(app):
    print("=== un artefacto de tinta pegado al borde (0 o h-1) de la franja no debe registrarse como numero real ===")
    rng = np.random.default_rng(12)
    franja, filas_y = _franja_sintetica(rng)
    franja[0:2, :] = 30  # simula el borde de la banda filtrandose en el recorte, como en la foto real
    obs = app._omr_detectar_number_y_obs(franja)
    assert not np.any(obs < 3.0), f"FALLO: un artefacto de borde se coló como observación real: {obs[obs < 3]}"
    print("OK\n")


def test_crosscheck_detecta_acuerdo(app):
    print("=== lattice de burbujas y de numeros EN FASE: crosscheck debe confirmar acuerdo (sin shift) ===")
    rng = np.random.default_rng(13)
    franja, filas_y = _franja_sintetica(rng)
    number_lat = {
        "y_centers": filas_y, "dy": DY, "y0": MARGEN,
        "number_alignment_confidence": 0.9, "n_obs": N_FILAS,
    }
    bubble_y_centers = filas_y + rng.normal(0, 0.5, size=N_FILAS)  # msima fase, jitter chico
    cc = app.omr_crosscheck_bubble_number(bubble_y_centers, number_lat, DY)
    print("row_crosscheck_error:", cc["row_crosscheck_error"])
    assert cc["row_crosscheck_error"] < 0.1
    assert cc["shift_confirmado"] is False
    print("OK\n")


def test_crosscheck_detecta_shift_de_una_fila(app):
    print("=== item 26: lattice de burbujas corrido +1 fila completa vs. el de numeros: crosscheck debe confirmarlo ===")
    rng = np.random.default_rng(14)
    franja, filas_y = _franja_sintetica(rng)
    number_lat = {
        "y_centers": filas_y, "dy": DY, "y0": MARGEN,
        "number_alignment_confidence": 0.9, "n_obs": N_FILAS,
    }
    bubble_y_centers_corrido = filas_y + DY  # toda la banda de burbujas corrida una fila completa
    cc = app.omr_crosscheck_bubble_number(bubble_y_centers_corrido, number_lat, DY)
    print("row_crosscheck_error:", cc["row_crosscheck_error"])
    assert cc["row_crosscheck_error"] >= 0.9
    assert cc["shift_confirmado"] is True
    print("OK\n")


def test_crosscheck_no_penaliza_number_strip_ilegible(app):
    print("=== item 26: number strip sin evidencia suficiente no debe poder tirar abajo geometry_confidence ===")
    rng = np.random.default_rng(15)
    filas_y = MARGEN + np.arange(N_FILAS) * DY
    number_lat_debil = {"y_centers": None, "dy": DY, "y0": None,
                         "number_alignment_confidence": 0.0, "n_obs": 2}
    cc = app.omr_crosscheck_bubble_number(filas_y, number_lat_debil, DY)
    assert cc["shift_confirmado"] is False
    assert cc["row_crosscheck_error"] is None
    print("OK\n")


def test_feature_flag_crosscheck_off_por_defecto(app):
    print("=== NUMBER_LATTICE_CROSSCHECK_HABILITADO debe estar OFF por defecto (falso positivo medido en foto real) ===")
    assert app.NUMBER_LATTICE_CROSSCHECK_HABILITADO is False, \
        "El veto del number lattice se habilitó sin evidencia de que no produce falsos positivos -- ver docstring del flag."
    print("OK\n")


if __name__ == "__main__":
    app = omr_metrics._cargar_app_module()
    test_number_lattice_recupera_periodo_y_origen(app)
    test_number_lattice_tolera_filas_faltantes(app)
    test_number_lattice_ignora_artefacto_de_borde(app)
    test_crosscheck_detecta_acuerdo(app)
    test_crosscheck_detecta_shift_de_una_fila(app)
    test_crosscheck_no_penaliza_number_strip_ilegible(app)
    test_feature_flag_crosscheck_off_por_defecto(app)
    print("TODO PASO - number lattice OK")
