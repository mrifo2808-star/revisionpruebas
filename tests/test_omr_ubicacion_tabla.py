"""
Tests de UBICACIÓN del bloque RESPUESTAS (localización de la tabla dentro de
la hoja completa) -- capa ANTERIOR al row lattice, encontrada rota con una
foto real reportada en producción: la app devolvía "no se detectó ninguna
respuesta" y los crops de revisión mostraban casillas de "N° DE FOLLETO" /
"IDENTIFICACIÓN DEL ESTUDIANTE" en vez de burbujas de RESPUESTAS.

Causa raíz (dos capas, ambas corregidas acá):

1. `omr_detectar_bloque_respuestas` cerraba el umbral binario con un kernel
   9x9 en 2 iteraciones (alcance efectivo ~16-18px), suficiente para puentear
   el hueco real medido entre secciones DISTINTAS de la hoja (13-14px:
   identificación / N° de folleto / RESPUESTAS) aunque mayor que el hueco
   DENTRO de una misma sección (2-6px) -- fusionaba varias secciones en un
   solo contorno gigante. Corregido: iterations=1.

2. Aun con las secciones separadas, `_omr_candidato_tiene_header` solo
   confirmaba "hay ALGUNA barra oscura sólida arriba" -- sin OCR no podía
   distinguir la barra real "RESPUESTAS" de la barra, también real, de
   "IDENTIFICACIÓN DEL ESTUDIANTE" o "CÉDULA DE IDENTIDAD" (que además es
   una grilla de círculos real -- dígitos 0-9 -- así que ni siquiera el
   fallback `_omr_candidato_evidencia_grilla`, que solo contaba círculos
   totales, la distinguía). Corregido exigiendo DENSIDAD DE FILAS real
   (`_omr_header_seguido_de_grilla` para la franja bajo el header,
   `_omr_candidato_evidencia_grilla` para el candidato completo): RESPUESTAS
   empaqueta hasta 20 filas de burbujas donde una sección de identificación
   tiene 1-3 filas de casillero de texto y CÉDULA tiene ~10-11 filas de
   dígitos -- una diferencia estructural amplia y medible sin leer texto.

Un tercer detalle encontrado en el camino: contar "corridas de bins por
encima de un piso fijo" en el histograma de posiciones Y no separa filas
reales adyacentes cuando el valle entre ellas no baja hasta ese piso (ruido
de detecciones Hough solapadas) -- fusionaba 5-6 filas reales en 1-3
"corridas". `_omr_contar_picos_locales` (máximos locales, no corridas)
corrige esto con un test unitario dedicado.

Correr desde la raíz del repo:  py -m pytest tests/test_omr_ubicacion_tabla.py -q
"""
import os
import sys

import cv2
import numpy as np

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import omr_metrics


def test_contar_picos_locales_separa_filas_con_ruido_entre_medio(app):
    print("=== _omr_contar_picos_locales: histograma real y ruidoso (5 filas reales) no debe "
          "fusionarse en menos filas solo porque el valle entre ellas no llega a cero ===")
    # Histograma real medido sobre una foto real (ver commit): 5 filas reales
    # de burbujas con ruido de detecciones Hough solapadas entre medio que
    # nunca baja del piso ingenuo (3), rompiendo un conteo por "corridas".
    hist = [0, 0, 0, 19, 1, 4, 5, 16, 3, 8, 7, 14, 3, 5, 8, 12, 8, 5, 2, 18, 0, 0]
    n_picos = app._omr_contar_picos_locales(hist, 3)
    print("picos encontrados:", n_picos)
    assert n_picos >= 4, f"FALLO: {n_picos} picos detectados, se esperaban al menos 4-5 filas reales"
    print("OK\n")


def test_contar_picos_locales_no_fabrica_picos_de_ruido_plano(app):
    print("=== _omr_contar_picos_locales: ruido bajo y parejo no debe contarse como filas ===")
    hist = [1, 2, 1, 2, 1, 1, 2, 1, 2, 1]
    n_picos = app._omr_contar_picos_locales(hist, 3)
    assert n_picos == 0, f"FALLO: se contaron {n_picos} picos en ruido plano por debajo del piso"
    print("OK\n")


def _imagen_seccion_identificacion(h=500, w=600):
    """Simula una sección de identificación: una barra de encabezado sólida
    seguida de 3 filas de CASILLERO DE TEXTO (rectángulos), sin ningún
    círculo real -- como "IDENTIFICACIÓN DEL ESTUDIANTE" en la plantilla."""
    img = np.full((h, w), 255, dtype=np.uint8)
    cv2.rectangle(img, (0, 0), (w, 30), 40, -1)  # barra de encabezado
    y = 60
    for _fila in range(3):
        for i in range(10):
            x0 = 20 + i * 55
            cv2.rectangle(img, (x0, y), (x0 + 45, y + 45), 80, 2)
        y += 90
    return img


def _imagen_seccion_respuestas(h=900, w=2400, n_filas=20, n_bandas=4, n_cols=5,
                                radio=6, spacing_x=26, spacing_y=20):
    """Simula una sección RESPUESTAS: barra de encabezado sólida seguida de
    n_filas x n_cols burbujas reales (círculos), repetido por banda."""
    img = np.full((h, w), 255, dtype=np.uint8)
    cv2.rectangle(img, (0, 0), (w, 25), 40, -1)
    x0 = 20
    for _banda in range(n_bandas):
        for row in range(n_filas):
            for col in range(n_cols):
                cx = x0 + col * spacing_x + 15
                cy = 50 + row * spacing_y + 15
                cv2.circle(img, (cx, cy), radio, 90, 2)
        x0 += n_cols * spacing_x + 40
    return img


def test_header_seguido_de_grilla_distingue_identificacion_de_respuestas(app):
    print("=== _omr_header_seguido_de_grilla: casillero de texto bajo un header real -- False; "
          "burbujas reales bajo un header real -- True ===")
    img_id = _imagen_seccion_identificacion()
    hb_id = app._omr_encontrar_barra_encabezado(img_id)
    assert hb_id > 0, "la barra de encabezado sintetica deberia detectarse"
    assert app._omr_header_seguido_de_grilla(img_id, hb_id) is False, \
        "un casillero de texto (sin circulos) no deberia pasar como grilla de burbujas"

    img_resp = _imagen_seccion_respuestas()
    hb_resp = app._omr_encontrar_barra_encabezado(img_resp)
    assert hb_resp > 0, "la barra de encabezado sintetica deberia detectarse"
    assert app._omr_header_seguido_de_grilla(img_resp, hb_resp) is True, \
        "una fila densa de burbujas reales deberia pasar como grilla de burbujas"
    print("OK\n")


def test_evidencia_grilla_rechaza_grilla_pequena_aunque_tenga_mas_circulos(app):
    print("=== _omr_candidato_evidencia_grilla: una grilla chica (tipo CEDULA, ~10 filas) con MAS "
          "circulos totales que una grilla RESPUESTAS real (~20 filas) NO debe aceptarse -- "
          "cuenta filas, no solo circulos crudos ===")
    # Grilla "tipo CEDULA": pocas filas (10) pero muchas columnas densas
    # (9), dando MAS circulos totales que una RESPUESTAS real mas angosta.
    img_cedula = np.full((700, 500), 255, dtype=np.uint8)
    cv2.rectangle(img_cedula, (0, 0), (500, 25), 40, -1)
    for row in range(10):
        for col in range(9):
            cx = 30 + col * 45
            cy = 60 + row * 40
            cv2.circle(img_cedula, (cx, cy), 9, 90, 2)
    rect_cedula = ((250.0, 350.0), (500.0, 700.0), 0.0)
    assert app._omr_candidato_evidencia_grilla(img_cedula, rect_cedula) is False, \
        "FALLO: una grilla de ~10 filas (tipo CEDULA) se acepto como si fuera RESPUESTAS"

    img_resp, rect_resp = None, None
    img = np.full((1000, 500), 255, dtype=np.uint8)
    cv2.rectangle(img, (0, 0), (500, 25), 40, -1)
    for row in range(20):
        for col in range(5):
            cx = 30 + col * 45
            cy = 60 + row * 34
            cv2.circle(img, (cx, cy), 9, 90, 2)
    rect_resp = ((225.0, 500.0), (450.0, 1000.0), 0.0)
    assert app._omr_candidato_evidencia_grilla(img, rect_resp) is True, \
        "una grilla real de 20 filas (RESPUESTAS) deberia aceptarse"
    print("OK\n")


if __name__ == "__main__":
    app = omr_metrics._cargar_app_module()
    test_contar_picos_locales_separa_filas_con_ruido_entre_medio(app)
    test_contar_picos_locales_no_fabrica_picos_de_ruido_plano(app)
    test_header_seguido_de_grilla_distingue_identificacion_de_respuestas(app)
    test_evidencia_grilla_rechaza_grilla_pequena_aunque_tenga_mas_circulos(app)
    print("TODO PASO - ubicacion de tabla RESPUESTAS OK")
