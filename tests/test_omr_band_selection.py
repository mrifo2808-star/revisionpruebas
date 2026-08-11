"""
Test de regresión permanente para un bug real encontrado con fotos reales
(smoke test privado, fuera del repo, 2026-08-11): cuando el bloque
RESPUESTAS detectado queda un poco ancho, el contenido vecino ("USO
EXCLUSIVO...") puede registrar suficiente tinta como para aparecer como un
candidato de banda MÁS ANCHO que las 4 columnas reales (que son angostas y
muy uniformes entre sí, mismo layout impreso). La selección de bandas previa
("si hay más candidatos que bandas esperadas, quedarse con los N más
anchos") descartaba entonces la columna real más angosta del grupo y
conservaba el bloque espurio -- resultado: una columna real completa (p.ej.
preguntas 41-60) nunca se leía, sus respuestas se leían en realidad desde la
columna vecina (61-80), y esa banda vecina quedaba corrida hacia el texto.
Esto es un desplazamiento SILENCIOSO de columna: mucho más grave que el caso
"USO EXCLUSIVO" ya cubierto en test_omr_regression.py, porque no
necesariamente cae por debajo del umbral de geometry_confidence -- las
burbujas de la columna vecina son reales y pueden leerse con confianza.

Este test reproduce el patrón de anchos exacto encontrado (4 candidatos
angostos ~60-71px + 1 candidato espurio ~165px) sobre una imagen sintética
(sin fotos reales, sin depender de ninguna en particular) y fija que la
selección de bandas se queda con las 4 columnas angostas/uniformes y
descarta el outlier ancho -- nunca al revés.

Correr desde la raíz del repo:  py tests/test_omr_band_selection.py
"""
import os
import sys

import cv2
import numpy as np

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import omr_metrics


def _gray_con_franjas(bw, bh, franjas, gris_fondo=255, gris_tinta=40):
    """franjas: lista de (x0, x1) -- cada una se pinta como una franja vertical
    oscura sólida, simulando la tinta de una columna real o de un bloque de
    texto vecino (a efectos de la proyección de tinta por columna, ambos
    lucen igual: una franja vertical densa)."""
    img = np.full((bh, bw), gris_fondo, dtype=np.uint8)
    for x0, x1 in franjas:
        img[:, x0:x1] = gris_tinta
    return img


def test_descarta_outlier_ancho_conserva_columnas_uniformes(app):
    print("=== 4 columnas angostas uniformes + 1 bloque ancho espurio: debe quedarse con las 4 uniformes ===")
    bw, bh = 480, 300
    # Mismo patron de anchos que el encontrado en fotos reales (ver docstring).
    franjas = [(17, 79), (91, 152), (164, 224), (236, 307), (315, 480)]
    gray = _gray_con_franjas(bw, bh, franjas)

    header_bottom, bands, bandas_fabricadas = app.omr_detectar_header_y_bandas(
        gray, max_bandas=4, bandas_esperadas=4, permitir_reparto_geometrico=True)

    print("bandas seleccionadas:", bands)
    assert len(bands) == 4, f"se esperaban 4 bandas, se obtuvieron {len(bands)}"

    anchos = [x1 - x0 for x0, x1 in bands]
    print("anchos:", anchos)
    # Las 4 columnas reales miden entre 60 y 71px -- el bloque espurio mide
    # 165px, un outlier claro. Ninguna banda seleccionada debe acercarse a eso.
    assert max(anchos) < 100, f"se conservó una banda demasiado ancha (probable bloque espurio): {anchos}"

    # El bloque espurio (315, 480) no debe estar representado por NINGUNA banda.
    for x0, x1 in bands:
        centro = (x0 + x1) / 2
        assert not (315 <= centro <= 480), \
            f"una banda quedó centrada dentro del bloque espurio (315-480): ({x0},{x1})"

    # Las 4 columnas reales SÍ deben estar todas representadas (una banda por
    # cada una, en orden, sin saltarse la más angosta).
    columnas_reales = [(17, 79), (91, 152), (164, 224), (236, 307)]
    for (rx0, rx1), (bx0, bx1) in zip(columnas_reales, bands):
        centro_real = (rx0 + rx1) / 2
        assert bx0 - 5 <= centro_real <= bx1 + 5, \
            f"la columna real ({rx0},{rx1}) no quedó cubierta por la banda seleccionada ({bx0},{bx1})"
    print("OK -- se conservaron las 4 columnas reales, se descartó el bloque espurio\n")


def test_sin_outlier_no_cambia_nada(app):
    print("=== control: con exactamente 4 columnas reales (sin candidato espurio) el resultado no cambia ===")
    bw, bh = 480, 300
    franjas = [(17, 79), (91, 152), (164, 224), (236, 307)]
    gray = _gray_con_franjas(bw, bh, franjas)
    header_bottom, bands, bandas_fabricadas = app.omr_detectar_header_y_bandas(
        gray, max_bandas=4, bandas_esperadas=4, permitir_reparto_geometrico=True)
    assert len(bands) == 4
    for (rx0, rx1), (bx0, bx1) in zip(franjas, bands):
        centro_real = (rx0 + rx1) / 2
        assert bx0 - 5 <= centro_real <= bx1 + 5
    print("OK -- comportamiento normal sin candidatos espurios no se altera\n")


if __name__ == "__main__":
    app = omr_metrics._cargar_app_module()
    test_descarta_outlier_ancho_conserva_columnas_uniformes(app)
    test_sin_outlier_no_cambia_nada(app)
    print("TODO PASO - seleccion de bandas OK")
