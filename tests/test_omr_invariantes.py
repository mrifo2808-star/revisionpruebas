"""
Tests sintéticos (sin fotos reales) de _omr_validar_invariantes_banda: los
invariantes geométricos duros que blindan contra grillas fabricadas o
corridas (orden de columnas/filas, spacing regular, ROIs dentro de la banda,
radio vs. spacing). No dependen de imágenes -- corren siempre, incluso sin
tests/data/omr/ poblado.

Incluye una regresión puntual: una versión anterior de este check exigía que
NINGÚN píxel del círculo de muestreo (radio incluido) sobrepasara el borde de
la banda ni por 1px, lo que marcaba como inválida CUALQUIER grilla real (las
burbujas de la fila/columna extrema tocan el borde por diseño). Se relajó
para exigir solo que el CENTRO esté dentro de la banda, con overshoot del
radio tolerado hasta un margen razonable -- este test fija ese comportamiento.

Correr desde la raíz del repo:  py tests/test_omr_invariantes.py
"""
import os
import sys

import numpy as np

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import omr_metrics


def _grilla_valida(bw=480, bh=940, n_filas=20, radio=6.0):
    x_centers = np.linspace(30, bw - 30, 5)
    y_centers = np.linspace(bh * 0.02, bh - bh * 0.02, n_filas)
    return x_centers, y_centers, radio, bw, bh, n_filas


def test_grilla_real_no_falla_por_borde(app):
    print("=== grilla real y regular: overshoot normal del radio en el borde NO debe penalizar ===")
    x_centers, y_centers, radio, bw, bh, n_filas = _grilla_valida()
    mult, violaciones = app._omr_validar_invariantes_banda(y_centers, x_centers, radio, bh, bw, n_filas)
    assert mult == 1.0, f"grilla regular real no debería perder confianza: mult={mult} viol={violaciones}"
    assert violaciones == [], f"no debería haber violaciones: {violaciones}"
    print("OK\n")


def test_columnas_desordenadas(app):
    print("=== columnas fuera de orden (xA<xB<...<xE roto) debe dar multiplicador 0 ===")
    x_centers, y_centers, radio, bw, bh, n_filas = _grilla_valida()
    x_centers = x_centers.copy()
    x_centers[2], x_centers[1] = x_centers[1], x_centers[2]  # rompe el orden estricto
    mult, violaciones = app._omr_validar_invariantes_banda(y_centers, x_centers, radio, bh, bw, n_filas)
    assert mult == 0.0
    assert "orden_columnas_A_E" in violaciones
    print("OK\n")


def test_centro_fuera_de_banda(app):
    print("=== un centro de columna fuera de los límites de la banda debe dar multiplicador 0 ===")
    x_centers, y_centers, radio, bw, bh, n_filas = _grilla_valida()
    x_centers = x_centers.copy()
    x_centers[-1] = bw + 50  # una "columna E" que cayó fuera de la banda
    mult, violaciones = app._omr_validar_invariantes_banda(y_centers, x_centers, radio, bh, bw, n_filas)
    assert mult == 0.0
    assert "centro_fuera_de_banda" in violaciones
    print("OK\n")


def test_spacing_horizontal_irregular(app):
    print("=== spacing horizontal muy irregular (columna colapsada) debe dar multiplicador 0 ===")
    x_centers, y_centers, radio, bw, bh, n_filas = _grilla_valida()
    x_centers = np.array([30.0, 32.0, 34.0, 36.0, 300.0])  # 4 casi pegadas + 1 lejos
    mult, violaciones = app._omr_validar_invariantes_banda(y_centers, x_centers, radio, bh, bw, n_filas)
    assert mult == 0.0
    assert "spacing_horizontal_irregular" in violaciones
    print("OK\n")


def test_radio_excede_spacing(app):
    print("=== radio de muestreo mayor que el spacing real entre burbujas debe degradar, no crashear ===")
    x_centers, y_centers, radio, bw, bh, n_filas = _grilla_valida()
    radio_enorme = float(np.diff(x_centers).min())  # radio == spacing completo -> se solapan
    mult, violaciones = app._omr_validar_invariantes_banda(y_centers, x_centers, radio_enorme, bh, bw, n_filas)
    assert mult < 1.0
    assert "radio_excede_spacing" in violaciones
    print("OK\n")


if __name__ == "__main__":
    app = omr_metrics._cargar_app_module()
    test_grilla_real_no_falla_por_borde(app)
    test_columnas_desordenadas(app)
    test_centro_fuera_de_banda(app)
    test_spacing_horizontal_irregular(app)
    test_radio_excede_spacing(app)
    print("TODO PASO - invariantes geometricos OK")
