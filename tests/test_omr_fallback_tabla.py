"""
Test sintético (sin fotos reales, sin OCR) del principio pedido en el cierre
de esta rama: un rectángulo grande y "sólido" NO puede convertirse en la
tabla RESPUESTAS solo por ser el candidato más grande -- necesita evidencia
real (header o un patrón denso de círculos), o `omr_detectar_bloque_respuestas`
debe devolver None (fail closed).

big_rectangle != answer_table sin evidencia estructural.

Antes del fix de esta iteración, la regla era "si ningún candidato tiene
header, usar igual el más grande" -- exactamente el mecanismo que produjo el
bug real "USO EXCLUSIVO" (ver test_omr_regression.py). Este test protege el
principio en general, con una imagen sintética que no depende de esa foto
puntual.

Correr desde la raíz del repo:  py tests/test_omr_fallback_tabla.py
"""
import os
import random
import sys

import cv2
import numpy as np

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import omr_metrics


def _imagen_bloque_texto_denso(h=1400, w=1000, seed=42):
    """Una hoja en blanco con un único bloque grande, sólido y apaisado (mismo
    fill/aspect/área que pasaría los umbrales de candidato a tabla RESPUESTAS)
    hecho de líneas horizontales oscuras muy juntas -- simula un bloque de
    texto denso o un logo, SIN header y SIN ningún patrón circular real."""
    rng = random.Random(seed)
    img = np.full((h, w), 255, dtype=np.uint8)
    x0, y0, x1, y1 = int(w * 0.55), int(h * 0.14), int(w * 0.95), int(h * 0.64)
    y = y0
    while y < y1:
        xlen = rng.randint(int((x1 - x0) * 0.4), int((x1 - x0) * 0.95))
        xstart = x0 + rng.randint(0, max(1, (x1 - x0) - xlen))
        cv2.line(img, (xstart, y), (xstart + xlen, y), 60, 3)
        y += 6
    return img


def _imagen_grilla_circulos_real(h=1400, w=1000, n_filas=20, n_cols=5, radio=8,
                                  spacing_x=30, spacing_y=34, x0=550, y0=200):
    """Una grilla real de círculos (como burbujas A-E x 20 filas), sin header --
    para confirmar que la evidencia geométrica SÍ es aceptada cuando existe de
    verdad (control positivo: la función no rechaza todo indiscriminadamente)."""
    img = np.full((h, w), 255, dtype=np.uint8)
    for row in range(n_filas):
        for col in range(n_cols):
            cx = x0 + col * spacing_x + 20
            cy = y0 + row * spacing_y + 20
            cv2.circle(img, (cx, cy), radio, 90, 2)
    rect = ((x0 + (n_cols * spacing_x) / 2, y0 + (n_filas * spacing_y) / 2),
            (n_cols * spacing_x + 40, n_filas * spacing_y + 40), 0.0)
    return img, rect


def test_bloque_texto_denso_no_se_acepta_como_tabla(app):
    print("=== bloque de texto denso, sin header ni patron OMR: NO debe convertirse en tabla RESPUESTAS ===")
    img = _imagen_bloque_texto_denso()
    img_bgr = cv2.cvtColor(img, cv2.COLOR_GRAY2BGR)
    quad = app.omr_detectar_bloque_respuestas(img_bgr)
    assert quad is None, ("FALLO CRITICO: un bloque grande sin evidencia real se aceptó como tabla "
                           "RESPUESTAS solo por ser el candidato más grande.")
    print("OK -- fail closed: sin header y sin evidencia de grilla, no hay tabla\n")


def test_evidencia_grilla_rechaza_texto_pero_acepta_circulos_reales(app):
    print("=== _omr_candidato_evidencia_grilla: control negativo (texto) y positivo (circulos reales) ===")
    img_texto = _imagen_bloque_texto_denso()
    rect_texto = ((775.0, 550.0), (400.0, 700.0), 0.0)  # cubre el bloque de texto denso
    assert app._omr_candidato_evidencia_grilla(img_texto, rect_texto) is False, \
        "texto denso sin circulos no deberia contar como evidencia de grilla"

    img_circulos, rect_circulos = _imagen_grilla_circulos_real()
    assert app._omr_candidato_evidencia_grilla(img_circulos, rect_circulos) is True, \
        "una grilla real de circulos SI deberia contar como evidencia (control positivo)"
    print("OK -- la funcion distingue texto denso de un patron real de burbujas, sin OCR\n")


if __name__ == "__main__":
    app = omr_metrics._cargar_app_module()
    test_bloque_texto_denso_no_se_acepta_como_tabla(app)
    test_evidencia_grilla_rechaza_texto_pero_acepta_circulos_reales(app)
    print("TODO PASO - fallback de tabla RESPUESTAS OK")
