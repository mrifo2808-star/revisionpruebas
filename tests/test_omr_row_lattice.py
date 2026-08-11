"""
Tests del ROW LATTICE (v4.1, ver omr_ajustar_lattice_vertical en app_revisor.py):
reemplaza la asignacion de los 20 centros de fila por kmeans 1D independiente
(que no sabe que las filas impresas son una estructura periodica y puede
acumular drift vertical dentro de una banda) por una grilla regular
y(i) = y0 + i*dy + delta_i, con y0/dy estimados de forma robusta y un
microajuste local acotado y suavizado.

Estos tests son SINTETICOS (sin fotos reales, no dependen de tests/data/omr/
poblado) porque no existe ground truth de posicion Y en pixeles para fotos
reales -- lo que si tenemos es control total sobre la verdad conocida al
construir observaciones Hough sinteticas. Miden explicitamente:

    row_error = abs(predicted_y - expected_y) / dy

con los objetivos del plan (item 19): mediana < 0.10 filas, P95 < 0.25,
maximo < 0.45 (nunca >= 0.5 -- una fila no puede "cruzar" a la posicion de
la siguiente). El item 20 (test de drift) se cubre midiendo el error en la
primera, la de mitad de banda y la ultima fila, y verificando que no crece
sistematicamente de punta a punta de la banda (row_drift_per_20 ~ 0).

Correr desde la raiz del repo:  py -m pytest tests/test_omr_row_lattice.py -q
"""
import os
import sys

import numpy as np

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import omr_metrics

N_FILAS = 20
BH = 900.0
BW = 480.0
# Mismo margen superior/inferior que usa _omr_y_centers_uniforme en produccion
# (bh*0.02) -- el anclaje de y0 en omr_ajustar_lattice_vertical (ver
# _omr_estimar_y0_robusto) asume ese mismo margen de plantilla como prior para
# desempatar ENTRE multiplos de dy, asi que el fixture sintetico tiene que
# modelar un margen realista, no uno arbitrario, o el "k" correcto puede
# quedar mas lejos del prior que un k vecino incorrecto (bug de test, no del
# algoritmo -- confirmado corriendo el caso a mano antes de este ajuste).
MARGEN_REAL = BH * 0.02
Y0_REAL = MARGEN_REAL
DY_REAL = (BH - 2 * MARGEN_REAL) / (N_FILAS - 1)
X_CENTERS_REAL = np.linspace(40, BW - 40, 5)


def _generar_obs(rng, y0=Y0_REAL, dy=DY_REAL, x_centers=X_CENTERS_REAL,
                  jitter_y=0.6, jitter_x=0.6, drop_frac=0.0, dy_drift_frac=0.0,
                  n_outliers=0):
    """Genera observaciones Hough sinteticas (x, y) para una banda con
    origen/periodo CONOCIDOS -- la "verdad" contra la que se mide row_error.

    dy_drift_frac: si > 0, aplica una perspectiva residual leve (el dy REAL
    crece linealmente a lo largo de la banda en vez de ser perfectamente
    constante) -- simula lo que una foto real de celular produce incluso
    despues del perspective-warp, para verificar que el microajuste local
    lo tolera sin producir el drift acumulado que el kmeans viejo si tenia.
    drop_frac: fraccion de filas completas (las 5 burbujas) que NO generan
    ninguna deteccion -- simula Hough perdiendo una fila por tinta debil.
    """
    # El paso tapera SIMETRICO alrededor de dy (crece hacia el final, decrece
    # hacia el principio) para que el pitch PROMEDIO se mantenga en dy -- una
    # version anterior lo hacia crecer monotonicamente desde dy, lo que sesga
    # el promedio hacia arriba y puede empujar la "verdad" sintetica mas alla
    # del propio borde inferior de la banda (bh) para drift_frac grande, algo
    # que una fila real jamas podria hacer (no hay tinta fuera de la banda).
    filas_y = []
    y_acumulado = y0
    for i in range(N_FILAS):
        filas_y.append(y_acumulado)
        paso = dy * (1.0 + dy_drift_frac * (i / max(1, N_FILAS - 1) - 0.5))
        y_acumulado += paso
    obs = []
    for i, y_fila in enumerate(filas_y):
        if rng.random() < drop_frac:
            continue
        for xc in x_centers:
            obs.append((xc + rng.normal(0, jitter_x), y_fila + rng.normal(0, jitter_y)))
    for _ in range(n_outliers):
        obs.append((rng.uniform(0, BW), rng.uniform(0, BH)))
    return np.array(obs), np.array(filas_y)


def _row_errors(y_centers, filas_y_verdad, dy):
    return np.abs(np.asarray(y_centers) - np.asarray(filas_y_verdad)) / dy


def test_grilla_regular_sin_ruido_recupera_filas_casi_exacto(app):
    print("=== grilla sintetica regular (sin jitter): row_error debe ser casi 0 en las 20 filas ===")
    rng = np.random.default_rng(1)
    obs, filas_y = _generar_obs(rng, jitter_y=0.0, jitter_x=0.0)
    lat = app.omr_ajustar_lattice_vertical(obs, X_CENTERS_REAL, BH, N_FILAS, radio_banda=6.0)
    err = _row_errors(lat["y_centers"], filas_y, lat["dy"])
    print("row_error:", np.round(err, 4))
    assert np.median(err) < 0.02
    assert np.max(err) < 0.05
    assert lat["row_alignment_confidence"] > 0.9
    print("OK\n")


def test_row_error_dentro_de_objetivos_con_jitter_realista(app):
    print("=== jitter realista (+/-0.6px) + perspectiva residual leve: row_error dentro de objetivos ===")
    rng = np.random.default_rng(2)
    obs, filas_y = _generar_obs(rng, jitter_y=0.6, jitter_x=0.6, dy_drift_frac=0.08)
    lat = app.omr_ajustar_lattice_vertical(obs, X_CENTERS_REAL, BH, N_FILAS, radio_banda=6.0)
    err = _row_errors(lat["y_centers"], filas_y, lat["dy"])
    print("median:", round(np.median(err), 4), "P95:", round(np.percentile(err, 95), 4), "max:", round(np.max(err), 4))
    assert np.median(err) < 0.10, f"mediana de row_error fuera de objetivo: {np.median(err)}"
    assert np.percentile(err, 95) < 0.25, f"P95 de row_error fuera de objetivo: {np.percentile(err, 95)}"
    assert np.max(err) < 0.45, f"maximo de row_error fuera de objetivo: {np.max(err)}"
    print("OK\n")


def test_row_offsets_nunca_cruzan_media_fila(app):
    print("=== el microajuste local (delta_i) nunca puede desplazar una fila a mitad de camino de la siguiente ===")
    rng = np.random.default_rng(3)
    obs, _filas_y = _generar_obs(rng, jitter_y=1.5, jitter_x=1.0, dy_drift_frac=0.15, n_outliers=15)
    lat = app.omr_ajustar_lattice_vertical(obs, X_CENTERS_REAL, BH, N_FILAS, radio_banda=6.0)
    max_delta_norm = np.max(np.abs(lat["row_offsets"])) / lat["dy"]
    print("max |delta_i|/dy:", round(max_delta_norm, 4))
    assert max_delta_norm < 0.5, f"FALLO CRITICO: un offset local cruzo la mitad del spacing: {max_delta_norm}"
    print("OK\n")


def test_no_hay_drift_sistematico_dentro_de_la_banda(app):
    print("=== item 20: caso tipico (perspectiva residual leve, ~3%) -- drift debe quedar cerca de cero ===")
    # Caso realista: despues del perspective-warp del pipeline, el dy real
    # deberia variar muy poco a lo largo de una banda (un par de % cuando
    # mucho) -- este es el caso que de verdad importa para producción.
    rng = np.random.default_rng(4)
    obs, filas_y = _generar_obs(rng, jitter_y=0.8, jitter_x=0.8, dy_drift_frac=0.03, n_outliers=8)
    lat = app.omr_ajustar_lattice_vertical(obs, X_CENTERS_REAL, BH, N_FILAS, radio_banda=6.0)
    err = _row_errors(lat["y_centers"], filas_y, lat["dy"])
    row_drift_per_20 = err[-1] - err[0]
    print(f"error P1={err[0]:.4f}  P10={err[N_FILAS // 2]:.4f}  P20={err[-1]:.4f}  row_drift_per_20={row_drift_per_20:.4f}")
    assert abs(row_drift_per_20) < 0.08, f"FALLO: drift sistematico en el caso tipico: {row_drift_per_20}"
    assert err[-1] < 0.15, f"la ultima fila del caso tipico no puede acumular error grande: {err[-1]}"
    print("OK\n")


def test_drift_acotado_incluso_en_escenario_adverso(app):
    print("=== item 20 (stress): incluso con perspectiva residual severa (10%, deliberadamente mas dura que un caso "
          "real tras el perspective-warp), el error debe quedar ACOTADO -- nunca la fuga sin limite que producia "
          "kmeans (donde el error de una fila podia crecer sin freno hasta ocupar la posicion de otra) ===")
    rng = np.random.default_rng(4)
    obs, filas_y = _generar_obs(rng, jitter_y=0.8, jitter_x=0.8, dy_drift_frac=0.10, n_outliers=8)
    lat = app.omr_ajustar_lattice_vertical(obs, X_CENTERS_REAL, BH, N_FILAS, radio_banda=6.0)
    err = _row_errors(lat["y_centers"], filas_y, lat["dy"])
    print(f"error P1={err[0]:.4f}  P10={err[N_FILAS // 2]:.4f}  P20={err[-1]:.4f}  max={err.max():.4f}")
    # El objetivo duro del plan (item 19/27): incluso en el peor caso nunca
    # se acepta como GEOMETRY_OK un desplazamiento >= 0.5 filas -- acá se
    # exige con margen (0.35) porque este escenario es deliberadamente mas
    # adverso que lo que deberia llegar al motor tras el perspective-warp.
    assert err.max() < 0.35, f"FALLO CRITICO: error sin acotar en escenario adverso: {err.max()}"
    print("OK\n")


def test_recupera_origen_aunque_falte_la_primera_fila(app):
    print("=== item 16: si la primera fila no genera ninguna deteccion, y0 no debe quedar corrido una fila completa ===")
    rng = np.random.default_rng(5)
    filas_y_completas = [Y0_REAL + i * DY_REAL for i in range(N_FILAS)]
    obs = []
    for i, y_fila in enumerate(filas_y_completas):
        if i == 0:
            continue  # fila 1 sin ninguna burbuja detectada (tinta debil / Hough la perdio)
        for xc in X_CENTERS_REAL:
            obs.append((xc + rng.normal(0, 0.5), y_fila + rng.normal(0, 0.5)))
    obs = np.array(obs)
    lat = app.omr_ajustar_lattice_vertical(obs, X_CENTERS_REAL, BH, N_FILAS, radio_banda=6.0)
    err0 = abs(lat["y_centers"][0] - filas_y_completas[0]) / lat["dy"]
    print("row_error fila 1 (sin evidencia propia):", round(err0, 4), " y0=", round(lat["y0"], 2),
          " esperado~", round(filas_y_completas[0], 2))
    assert err0 < 0.35, f"FALLO: el origen quedo corrido sin evidencia de la primera fila (ROW_INDEX_SHIFT): {err0}"
    print("OK\n")


def test_dy_global_regulariza_banda_con_evidencia_debil(app):
    print("=== item 15: una banda con pocas detecciones debe apoyarse en dy_global en vez de caer a UNIFORM_FALLBACK puro ===")
    rng = np.random.default_rng(6)
    # Banda con evidencia muy escasa (borrosa) pero real: unas pocas filas.
    filas_y_completas = [Y0_REAL + i * DY_REAL for i in range(N_FILAS)]
    obs = []
    for i in (3, 4, 9, 10, 15, 16):
        for xc in X_CENTERS_REAL:
            obs.append((xc + rng.normal(0, 0.5), filas_y_completas[i] + rng.normal(0, 0.5)))
    obs = np.array(obs)
    dy_global = DY_REAL * 1.0  # ya demostrado por las otras 3 bandas hermanas
    lat_con_prior = app.omr_ajustar_lattice_vertical(obs, X_CENTERS_REAL, BH, N_FILAS, radio_banda=6.0, dy_prior=dy_global)
    err = _row_errors(lat_con_prior["y_centers"], filas_y_completas, lat_con_prior["dy"])
    print("dy encontrado:", round(lat_con_prior["dy"], 3), " dy_global:", round(dy_global, 3))
    print("median row_error con dy_global:", round(np.median(err), 4))
    assert abs(lat_con_prior["dy"] - dy_global) / dy_global < 0.12
    assert np.median(err) < 0.15
    print("OK\n")


def test_banda_sin_evidencia_cae_a_uniform_fallback(app):
    print("=== banda genuinamente vacia (foto no llego a mostrar esa columna): debe fallar cerrado, no fabricar filas ===")
    lat = app.omr_ajustar_lattice_vertical(np.empty((0, 2)), X_CENTERS_REAL, BH, N_FILAS, radio_banda=6.0)
    assert lat["row_source"] == "UNIFORM_FALLBACK"
    assert lat["row_alignment_confidence"] == 0.0
    assert len(lat["y_centers"]) == N_FILAS
    print("OK\n")


if __name__ == "__main__":
    app = omr_metrics._cargar_app_module()
    test_grilla_regular_sin_ruido_recupera_filas_casi_exacto(app)
    test_row_error_dentro_de_objetivos_con_jitter_realista(app)
    test_row_offsets_nunca_cruzan_media_fila(app)
    test_no_hay_drift_sistematico_dentro_de_la_banda(app)
    test_drift_acotado_incluso_en_escenario_adverso(app)
    test_recupera_origen_aunque_falte_la_primera_fila(app)
    test_dy_global_regulariza_banda_con_evidencia_debil(app)
    test_banda_sin_evidencia_cae_a_uniform_fallback(app)
    print("TODO PASO - row lattice OK")
