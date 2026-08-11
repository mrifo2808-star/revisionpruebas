"""
Motor OMR (Optical Mark Recognition) para la plantilla "PLANTILLA DE HOJA DE
RESPUESTAS" — detecta automaticamente que alternativa (A-E) marco el estudiante
en cada fila de la tabla RESPUESTAS. Este motor es la ÚNICA fuente de las
respuestas: Claude Vision NUNCA lee ni confirma burbujas (la app la usa solo
para transcribir nombre/RUT de la cabecera, una tarea completamente aparte).
Las preguntas que el motor no logra determinar con confianza quedan como
"dudosas" para que la persona las revise a mano en la app -- ninguna IA
adivina ni corrige una lectura de burbujas.

Diseño: en vez de coordenadas de pixel fijas (fragil ante fotos de celular con
angulos/distancias distintas), todo se deriva de la propia imagen:
  1. localizar el bloque RESPUESTAS (umbral adaptativo + contornos: el bloque es
     el mayor recuadro "solido" del documento) -- solo si la imagen es la hoja
     completa; si ya viene recortada (modo "solo_respuestas" o las mitades
     izq/der que genera la propia app) se usa la imagen entera directamente;
  2. si se localizo, enderezarlo con una transformacion de perspectiva a partir
     de sus 4 esquinas;
  3. ubicar la barra de encabezado (si existe) y las bandas de columnas (1 a 4,
     segun cuanto de la tabla se ve) por proyeccion de tinta -- probando varios
     umbrales de sensibilidad, y si ninguno logra distinguir suficientes huecos
     reales entre columnas (foto borrosa, poco contraste, muy comprimida),
     repartiendo el ancho de contenido en partes iguales en vez de fallar (la
     plantilla tiene columnas de ancho fijo por diseño). Esto nunca reescala ni
     comprime la imagen: solo cambia dónde se traza el límite de cada columna;
  4. ajustar, PARA CADA BANDA POR SEPARADO, una grilla de 20 filas x 5 columnas
     usando deteccion de circulos (Hough) + kmeans 1D -- robusta a que algunas
     burbujas no tengan ningun circulo detectable; si la foto es muy borrosa o
     pequeña y Hough no encuentra suficientes circulos en alguna banda puntual,
     esa banda cae a una grilla uniforme por proporciones (menos precisa pero
     nunca falla) sin afectar a las demás bandas. Las filas NO se comparten
     entre bandas a propósito: en pruebas con fotos reales, un ajuste global
     compartido llegó a acumular un corrimiento de una fila completa dentro de
     una sola columna, produciendo respuestas incorrectas con ALTA confianza en
     vez de quedar marcadas como dudosas -- el peor escenario posible. Ajustar
     cada banda por separado evita que un corrimiento en una columna contamine
     a las demás;
  5. medir cuanto mas oscura esta cada burbuja que el papel en blanco, comparando
     SIEMPRE dentro de la misma fila (igual que la regla que ya usa el prompt de
     Claude), y clasificar con umbrales configurables.

Validado contra fotos reales de celular (no plantillas escaneadas limpias): además
de identificar correctamente marcas fuertes, marcas débiles (con confianza más
baja) y preguntas dejadas en blanco, este ajuste por banda se agregó DESPUÉS de
encontrar mediante verificación manual (comparando contra la foto pregunta por
pregunta) que la grilla global previa podía desalinearse una fila completa a
partir de cierto punto de una columna -- ver commit que introdujo esta versión
para el detalle de cómo se detectó y qué caso real lo confirmó.
"""

import numpy as np
import cv2


# ─── Umbrales configurables (ajustables durante pruebas, no hardcodeados en la
# logica de clasificacion) ────────────────────────────────────────────────────
# Todos son relativos a una escala normalizada 0..1 = (oscuridad - oscuridad
# tipica del papel en blanco) / (oscuridad tipica de una marca - papel en blanco).
THRESHOLDS = {
    # confianza minima (norm_best) para considerar que SI hay una marca real
    "MIN_MARK_SCORE": 0.25,
    # margen minimo entre la mejor y la segunda mejor alternativa para no
    # considerarlo ambiguo/empate
    "AMBIGUOUS_MARGIN": 0.15,
    # margen a partir del cual se considera alta confianza (se usa directo, sin
    # marcar la pregunta como dudosa)
    "HIGH_CONFIDENCE_MARGIN": 0.30,
    # si dos alternativas superan este nivel de oscurecimiento ambas, se
    # considera "doble marca" aunque el margen entre ellas sea alto
    "DOUBLE_MARK_THRESHOLD": 0.55,
    # una pregunta "sin_marca" con omr_confidence por ENCIMA de este valor (pero
    # aun por debajo de MIN_MARK_SCORE) se trata como "posible marca muy débil"
    # y se marca igual como dudosa para revisión manual, en vez de asumir en
    # silencio que está en blanco
    "BLANK_REVIEW_MARGIN": 0.12,
}

N_FILAS_POR_BLOQUE = 20  # fijo por diseño de la plantilla impresa (igual que usa el prompt de Claude)
LETRAS = ["A", "B", "C", "D", "E"]


class OMRError(Exception):
    """Fallo irrecuperable de una etapa del pipeline OMR — la hoja queda sin
    resolver para completar a mano (nunca se reemplaza con una lectura de IA)."""


# ─── 1) Localizacion del bloque RESPUESTAS dentro de una hoja completa ───────

def detectar_bloque_respuestas(img_bgr: np.ndarray):
    """
    Devuelve los 4 puntos (rotados, en cualquier orden) del rectangulo que
    contiene la tabla RESPUESTAS dentro de una foto de la hoja completa, o None
    si no se encuentra con confianza suficiente (la hoja completa tiene ademas
    los bloques de identificacion/cedula, que son mas chicos y con menor
    fill_ratio que el bloque RESPUESTAS, el mayor recuadro "solido" del diseño).
    """
    gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)
    h, w = gray.shape
    th = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                cv2.THRESH_BINARY_INV, 35, 15)
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (9, 9))
    closed = cv2.morphologyEx(th, cv2.MORPH_CLOSE, kernel, iterations=2)
    contours, _ = cv2.findContours(closed, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)

    best, best_area = None, -1
    page_area = h * w
    for c in contours:
        area = cv2.contourArea(c)
        if area < page_area * 0.03:  # descarta ruido pequeño
            continue
        rect = cv2.minAreaRect(c)
        (_, _), (rw, rh), _ = rect
        if rw < 1 or rh < 1:
            continue
        fill = area / (rw * rh)
        long_side, short_side = max(rw, rh), min(rw, rh)
        aspect = long_side / max(short_side, 1e-6)
        # el bloque RESPUESTAS es solido (fill alto) y moderadamente apaisado
        # (4 columnas una al lado de otra); evita capturar la hoja entera
        # (aspect muy alto / area ~ toda la pagina) o cajas chicas de cedula/ID.
        if fill > 0.55 and 1.0 < aspect < 2.6 and area < page_area * 0.6:
            if area > best_area:
                best_area, best = area, rect
    if best is None:
        return None
    return cv2.boxPoints(best)


def _ordenar_puntos(pts: np.ndarray) -> np.ndarray:
    pts = np.array(pts)
    s = pts.sum(axis=1)
    d = np.diff(pts, axis=1).flatten()
    tl = pts[np.argmin(s)]
    br = pts[np.argmax(s)]
    tr = pts[np.argmin(d)]
    bl = pts[np.argmax(d)]
    return np.array([tl, tr, br, bl], dtype="float32")


def enderezar_region(img_bgr: np.ndarray, quad: np.ndarray) -> np.ndarray:
    """Perspective-warp de un cuadrilatero (4 puntos, cualquier orden/rotacion) a un rectangulo recto."""
    src = _ordenar_puntos(quad)
    w_top = np.linalg.norm(src[1] - src[0])
    w_bot = np.linalg.norm(src[2] - src[3])
    h_left = np.linalg.norm(src[3] - src[0])
    h_right = np.linalg.norm(src[2] - src[1])
    out_w = int(max(w_top, w_bot))
    out_h = int(max(h_left, h_right))
    dst = np.array([[0, 0], [out_w - 1, 0], [out_w - 1, out_h - 1], [0, out_h - 1]], dtype="float32")
    M = cv2.getPerspectiveTransform(src, dst)
    return cv2.warpPerspective(img_bgr, M, (out_w, out_h))


# ─── 2) Barra de encabezado + bandas de columnas ─────────────────────────────

def _bandas_por_gaps(col_smooth: np.ndarray, bw: int, umbral: float):
    mask = col_smooth > umbral
    bands = []
    in_band, start = False, 0
    for x in range(bw):
        if mask[x] and not in_band:
            in_band, start = True, x
        elif not mask[x] and in_band:
            in_band = False
            bands.append((start, x))
    if in_band:
        bands.append((start, bw))
    bands = [b for b in bands if (b[1] - b[0]) > bw * 0.04]
    return bands, mask


def detectar_header_y_bandas(gray: np.ndarray, max_bandas: int = 4, bandas_esperadas: int = None,
                              permitir_reparto_geometrico: bool = True):
    """
    Devuelve (header_bottom, bandas) donde bandas es una lista de (x0,x1) — una
    por cada bloque de 20 preguntas visible en la imagen (1 a 4, segun si la
    imagen es la hoja/bloque completo o un recorte parcial). header_bottom=0 si
    no se detecta una barra de encabezado solida (caso de imagenes ya recortadas
    que empiezan directo en la fila 1).

    La separacion de columnas se basa en encontrar huecos en blanco entre ellas
    (proyeccion de tinta), pero eso puede fallar en fotos borrosas, con poca luz
    o muy comprimidas (p.ej. reenviadas por WhatsApp): el hueco entre columnas
    puede no leerse lo bastante "blanco" como para distinguirse. Por eso:
      1. se prueban varios umbrales de sensibilidad, del mas estricto al mas
         permisivo, hasta encontrar `bandas_esperadas` bandas reales;
      2. si ninguno lo logra Y `permitir_reparto_geometrico` es True, se cae a
         dividir el ancho total de contenido en `bandas_esperadas` franjas
         iguales -- la plantilla tiene columnas de ancho fijo por diseño, asi
         que un reparto proporcional es una base razonable incluso sin poder
         ver los huecos con claridad. Esto NUNCA implica reescalar ni
         comprimir la imagen: se sigue trabajando sobre los mismos pixeles
         originales, solo cambia dónde se traza el límite de cada columna.

    IMPORTANTE sobre `permitir_reparto_geometrico`: este reparto asume que el
    contenido visible SÍ contiene las `bandas_esperadas` columnas completas,
    solo que sus huecos no se distinguen -- una suposicion razonable cuando la
    región ya fue validada de forma independiente como la tabla RESPUESTAS
    completa (detectar_bloque_respuestas). Si en cambio la imagen es un
    recorte arbitrario que el usuario subió (es_recorte=True) y genuinamente
    NO alcanza a mostrar todas las columnas, forzar el reparto fabricaría una
    columna que no existe y devolvería respuestas con confianza alta pero
    incorrectas -- el peor escenario posible. Por eso analizar_imagen() solo
    pasa permitir_reparto_geometrico=True cuando ya validó la región completa.

    `bandas_esperadas` normalmente es ceil(n_preguntas/20) -- cuantas columnas
    de 20 preguntas hacen falta para cubrir la pauta configurada.
    """
    h, w = gray.shape
    row_mean = gray.mean(axis=1)
    dark_rows = row_mean < (row_mean.mean() - row_mean.std() * 0.5)
    header_bottom = 0
    in_run, run_start = False, 0
    for y in range(min(h, int(h * 0.25))):
        if dark_rows[y] and not in_run:
            in_run, run_start = True, y
        elif not dark_rows[y] and in_run:
            in_run = False
            if (y - run_start) > h * 0.015:
                header_bottom = y
    body = gray[header_bottom:h, :]
    bh, bw = body.shape

    col_ink = 255 - body.astype(np.float32).mean(axis=0)
    rng = col_ink.max() - col_ink.min()
    if rng < 8:
        # Contraste casi nulo de lado a lado: no hay columnas de texto/burbujas
        # reales que distinguir (p.ej. una foto de otra cosa, o una hoja en
        # blanco). El reparto proporcional de más abajo asume que SÍ hay una
        # tabla real por leer, así que mejor fallar acá con un mensaje claro
        # que "leer" una tabla inexistente con confianza inventada.
        raise OMRError("La imagen no tiene suficiente contraste para ubicar columnas "
                        "(¿es realmente una foto de la tabla RESPUESTAS?).")
    col_norm = (col_ink - col_ink.min()) / rng
    k = np.ones(5) / 5
    col_smooth = np.convolve(col_norm, k, mode="same")

    objetivo = bandas_esperadas or max_bandas
    bands, mask = [], col_smooth > 0.15
    for umbral in (0.15, 0.10, 0.07, 0.045, 0.03):
        candidatas, mask_u = _bandas_por_gaps(col_smooth, bw, umbral)
        if len(candidatas) >= objetivo:
            bands, mask = candidatas, mask_u
            break
        if len(candidatas) > len(bands):
            bands, mask = candidatas, mask_u  # nos quedamos con el mejor intento aunque no alcance el objetivo

    if len(bands) > max_bandas:
        bands = sorted(bands, key=lambda b: b[1] - b[0], reverse=True)[:max_bandas]
        bands = sorted(bands, key=lambda b: b[0])

    if len(bands) < objetivo:
        if not permitir_reparto_geometrico:
            raise OMRError(
                f"Solo se distinguen {len(bands)} de las {objetivo} columnas esperadas y esta imagen "
                "es un recorte sin validar (no la hoja completa) -- no se fuerza un reparto geométrico "
                "porque podría fabricar una columna que no está realmente en la foto.")
        # Ningún umbral encontró suficientes huecos reales entre columnas (foto
        # borrosa/comprimida/poco contraste): repartir el ancho de contenido en
        # partes iguales en vez de rendirse. Se usa el rango donde SÍ hay algo
        # de tinta (no el ancho completo de la imagen) para no incluir márgenes
        # en blanco de los bordes en el reparto.
        idx_contenido = np.where(col_smooth > 0.03)[0]
        x0c, x1c = (int(idx_contenido[0]), int(idx_contenido[-1]) + 1) if len(idx_contenido) else (0, bw)
        ancho = (x1c - x0c) / objetivo
        bands = [(int(x0c + i * ancho), int(x0c + (i + 1) * ancho)) for i in range(objetivo)]

    return header_bottom, bands


# ─── 3) Ajuste de grilla 20x5 por banda ──────────────────────────────────────

def _kmeans_1d(values: np.ndarray, k: int, iters: int = 50) -> np.ndarray:
    values = np.asarray(values, dtype=np.float64)
    lo, hi = values.min(), values.max()
    centers = np.linspace(lo, hi, k) if hi > lo else np.full(k, lo)
    for _ in range(iters):
        d = np.abs(values[:, None] - centers[None, :])
        assign = d.argmin(axis=1)
        new_centers = centers.copy()
        for i in range(k):
            pts = values[assign == i]
            if len(pts) > 0:
                new_centers[i] = pts.mean()
        if np.allclose(new_centers, centers):
            break
        centers = new_centers
    return np.sort(centers)


def _y_centers_uniforme(bh: float, n_filas: int) -> np.ndarray:
    margen_y = bh * 0.02
    return np.linspace(margen_y, bh - margen_y, n_filas)


def ajustar_grilla(body_gray: np.ndarray, bands: list, n_filas: int = N_FILAS_POR_BLOQUE):
    """
    Ajusta, PARA CADA BANDA POR SEPARADO, sus 20 centros de fila y sus 5 centros
    de columna, usando deteccion de circulos + kmeans 1D.

    IMPORTANTE: los centros de fila NO se comparten entre bandas aunque en la
    plantilla las filas estén alineadas horizontalmente entre columnas. Se
    intentó eso primero y falló en la práctica: en fotos reales de celular el
    ajuste global por kmeans puede acumular un corrimiento de hasta una fila
    completa dentro de una sola banda (columna) sin que ninguna otra banda se
    vea afectada -- eso produce respuestas incorrectas con ALTA confianza en
    vez de quedar marcadas como dudosas, que es el peor escenario posible. Al
    ajustar cada banda de forma independiente, un corrimiento en una columna no
    contamina a las demás, y dentro de esa columna el ajuste sigue siendo
    consistente porque usa solo los círculos que realmente están en esa franja.

    Si Hough no encuentra suficientes círculos en alguna banda puntual (foto
    borrosa/pequeña), esa banda específica cae a una grilla uniforme por
    proporciones -- las demás bandas siguen usando su ajuste por círculos.
    """
    bh, bw = body_gray.shape
    blur = cv2.medianBlur(body_gray, 3)
    min_band_w = min(x1 - x0 for x0, x1 in bands)
    r_guess = max(3, min_band_w // 12)
    circles = cv2.HoughCircles(
        blur, cv2.HOUGH_GRADIENT, dp=1, minDist=max(4, r_guess),
        param1=60, param2=14, minRadius=max(2, int(r_guess * 0.6)),
        maxRadius=int(r_guess * 1.8),
    )
    pts = circles[0] if circles is not None else np.empty((0, 3))

    y_centers_por_banda, band_x_centers, radios = [], [], []
    for (x0, x1) in bands:
        sel = pts[(pts[:, 0] >= x0) & (pts[:, 0] < x1)]
        if len(sel) >= max(5 * n_filas * 0.35, 10):
            y_centers_por_banda.append(_kmeans_1d(sel[:, 1], n_filas))
            band_x_centers.append(_kmeans_1d(sel[:, 0], 5))
            radios.append(float(np.median(sel[:, 2])))
        else:
            # Esta banda puntual no tiene suficientes círculos detectados:
            # grilla uniforme por proporciones solo para ella.
            y_centers_por_banda.append(_y_centers_uniforme(bh, n_filas))
            bw_band = x1 - x0
            num_w = bw_band * 0.18  # ~18% inicial para el número de pregunta impreso
            xs = x0 + num_w + (np.arange(5) + 0.5) * ((bw_band - num_w) / 5)
            band_x_centers.append(xs)
            radios.append(max(3.0, min_band_w * 0.08))

    radio = float(np.median(radios))
    return np.array(y_centers_por_banda), np.array(band_x_centers), radio


# ─── 4) Puntaje de relleno + clasificacion ───────────────────────────────────

def _oscuridad_celda(gray: np.ndarray, cx: float, cy: float, r: float) -> float:
    x0, x1 = max(0, int(cx - r)), int(cx + r + 1)
    y0, y1 = max(0, int(cy - r)), int(cy + r + 1)
    patch = gray[y0:y1, x0:x1]
    if patch.size == 0:
        return 0.0
    return 255.0 - float(patch.mean())


def calcular_puntajes(body_gray: np.ndarray, y_centers_por_banda, band_x_centers, radio: float):
    """Devuelve lista de dicts {A:.., B:.., ..} de oscuridad cruda por pregunta, en orden P1..Pn.
    y_centers_por_banda: array (n_bandas, n_filas) -- una grilla de filas independiente por banda."""
    sample_r = max(2.0, radio * 0.75)
    filas = []
    for bi, band_x in enumerate(band_x_centers):
        for cy in y_centers_por_banda[bi]:
            fila = {LETRAS[li]: _oscuridad_celda(body_gray, cx, cy, sample_r)
                    for li, cx in enumerate(band_x)}
            filas.append(fila)
    return filas


def clasificar_pregunta(scores: dict, baseline: float, peak: float, umbrales: dict = THRESHOLDS) -> dict:
    letras = list(scores.keys())
    vals = np.array([scores[l] for l in letras])
    rng = max(peak - baseline, 1e-6)
    norm = (vals - baseline) / rng
    order = np.argsort(norm)[::-1]
    best_i, second_i = order[0], order[1]
    best_v, second_v = norm[best_i], norm[second_i]
    margin = best_v - second_v

    doble = best_v > umbrales["DOUBLE_MARK_THRESHOLD"] and second_v > umbrales["DOUBLE_MARK_THRESHOLD"]

    if best_v < umbrales["MIN_MARK_SCORE"]:
        return {"letra": None, "status": "sin_marca",
                "omr_confidence": round(max(0.0, min(1.0, float(best_v))), 3)}
    if doble:
        return {"letra": None, "status": "doble_marca",
                "omr_confidence": round(float(margin), 3)}
    if margin < umbrales["AMBIGUOUS_MARGIN"]:
        return {"letra": None, "status": "ambiguo",
                "omr_confidence": round(float(margin), 3)}
    alta = margin >= umbrales["HIGH_CONFIDENCE_MARGIN"]
    conf = max(0.0, min(1.0, 0.5 + margin))
    return {"letra": letras[best_i],
            "status": "alta_confianza" if alta else "confianza_media",
            "omr_confidence": round(float(conf), 3)}


# ─── 5) Recorte de una pregunta puntual (para revision manual) ──────────────

def recortar_pregunta(body_bgr: np.ndarray, y_centers_por_banda, band_x_centers, radio: float,
                       idx_local: int, n_bandas: int) -> np.ndarray:
    """
    Recorta la fila completa (5 burbujas + algo de contexto) de la pregunta con
    indice local idx_local (0-based, orden banda por banda como en calcular_puntajes),
    directamente sobre la imagen SIN comprimir, para mostrársela a la persona
    en la UI de corrección manual. Incluye holgura vertical/horizontal generosa
    por si la grilla esta levemente desalineada.
    """
    n_filas = len(y_centers_por_banda[0])
    bi, ri = divmod(idx_local, n_filas)
    if bi >= len(band_x_centers):
        raise OMRError(f"idx_local={idx_local} cae fuera de las {len(band_x_centers)} bandas detectadas "
                        f"({n_filas} filas/banda) — la grilla no cubre esa pregunta.")
    band_x = band_x_centers[bi]
    cy = y_centers_por_banda[bi][ri]
    x0 = band_x[0] - radio * 3.2
    x1 = band_x[-1] + radio * 2.2
    y0 = cy - radio * 2.4
    y1 = cy + radio * 2.4
    h, w = body_bgr.shape[:2]
    x0, x1 = max(0, int(x0)), min(w, int(x1))
    y0, y1 = max(0, int(y0)), min(h, int(y1))
    return body_bgr[y0:y1, x0:x1]


# ─── 6) Diagnostico visual ────────────────────────────────────────────────────

COLOR_ALTA = (60, 180, 60)      # verde (BGR)
COLOR_MEDIA = (0, 200, 230)     # amarillo/naranjo
COLOR_AMBIGUA = (40, 40, 220)   # rojo
COLOR_BLANCO = (150, 150, 150)  # gris (sin marca, con confianza suficiente para no ser dudosa)

def anotar_diagnostico(body_bgr: np.ndarray, y_centers_por_banda, band_x_centers, radio: float,
                        resultados: list, offset_pregunta: int = 0, escala: int = 3) -> np.ndarray:
    """Genera la imagen de diagnostico: un circulo alrededor de la burbuja elegida
    (o un punto junto a la fila si no se eligio ninguna) y una etiqueta "Pn:letra"
    coloreada segun el estado final de cada pregunta. `escala` agranda la imagen
    para que las etiquetas sean legibles en pantalla."""
    base = cv2.resize(body_bgr, (body_bgr.shape[1] * escala, body_bgr.shape[0] * escala),
                       interpolation=cv2.INTER_NEAREST)
    dbg = base.copy()
    n_filas = len(y_centers_por_banda[0])
    for idx, r in enumerate(resultados):
        bi, ri = divmod(idx, n_filas)
        if bi >= len(band_x_centers):
            continue  # pregunta sin banda/fila real en esta grilla (no debería llegar aquí; se ignora en vez de romper el diagnóstico)
        band_x = band_x_centers[bi]
        cy = y_centers_por_banda[bi][ri] * escala
        q = offset_pregunta + idx + 1
        status = r["status"]
        color = {
            # estados crudos del clasificador OMR (uso directo del motor)
            "alta_confianza": COLOR_ALTA, "confianza_media": COLOR_MEDIA, "sin_marca": COLOR_BLANCO,
            # metodo_por_pregunta del pipeline en app_revisor.py (todo resuelto por OMR, nunca por IA)
            "confiable": COLOR_ALTA, "revisar_media": COLOR_MEDIA,
            "revisar_dudoso": COLOR_AMBIGUA, "blanco": COLOR_BLANCO,
        }.get(status, COLOR_AMBIGUA)
        if r["letra"]:
            li = LETRAS.index(r["letra"])
            cx = band_x[li] * escala
            cv2.circle(dbg, (int(cx), int(cy)), int(radio * 1.3 * escala), color, max(1, escala // 2))
        else:
            cv2.circle(dbg, (int(band_x[0] * escala - radio * escala), int(cy)), 3 * escala, color, -1)
        conf_pct = round(r.get("omr_confidence", 0) * 100)
        label = f"{q}:{r['letra'] or '?'} {conf_pct}%"
        cv2.putText(dbg, label, (int(band_x[0] * escala - radio * 3.4 * escala), int(cy) - int(radio * escala) - 2),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.32 * escala / 2, color, 1, cv2.LINE_AA)
    return dbg


# ─── 7) Orquestacion de una imagen (hoja completa o recorte) ────────────────

def analizar_imagen(img_bgr: np.ndarray, es_recorte: bool, max_bandas: int = 4, n_preguntas: int = None) -> dict:
    """
    Punto de entrada principal: toma una imagen ya cargada (BGR, resolucion
    original, SIN recomprimir) y devuelve la grilla ajustada + resultados OMR
    crudos por pregunta local (orden banda por banda, 20 filas por banda).

    es_recorte=False: se asume que es la hoja completa (o al menos incluye la
      cabecera) -> se localiza el bloque RESPUESTAS y se endereza antes de leer
      la grilla.
    es_recorte=True: se asume que la imagen YA es (una porcion de) la tabla
      RESPUESTAS -> no se busca ni se recorta nada, se trabaja directo sobre
      toda la imagen (con o sin barra de encabezado visible).
    n_preguntas: si se indica, cuantas bandas (columnas de 20) hacen falta para
      cubrirlo -- permite que detectar_header_y_bandas caiga a un reparto
      geometrico proporcional cuando la foto no tiene contraste suficiente
      para distinguir los huecos reales entre columnas, en vez de fallar.

    Lanza OMRError si no logra ubicar una grilla legible (la hoja queda sin
    resolver para completar a mano; nunca se reemplaza con una lectura de IA).
    """
    img = img_bgr
    if not es_recorte:
        quad = detectar_bloque_respuestas(img)
        if quad is None:
            raise OMRError("No se pudo localizar el bloque RESPUESTAS en la hoja completa.")
        img = enderezar_region(img, quad)

    bandas_esperadas = min(max_bandas, -(-n_preguntas // N_FILAS_POR_BLOQUE)) if n_preguntas else None
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    # El reparto geométrico de columnas (ver detectar_header_y_bandas) solo se
    # permite cuando la región ya fue validada como la tabla RESPUESTAS
    # completa (es_recorte=False, arriba). Para un recorte que subió el
    # usuario directamente, no hay forma de confirmar que de verdad contiene
    # todas las columnas esperadas, así que ante la duda se prefiere fallar
    # con un mensaje claro antes que fabricar una columna inexistente.
    header_bottom, bands = detectar_header_y_bandas(
        gray, max_bandas=max_bandas, bandas_esperadas=bandas_esperadas,
        permitir_reparto_geometrico=not es_recorte)
    body_gray = gray[header_bottom:, :]
    body_bgr = img[header_bottom:, :]

    y_centers_por_banda, band_x_centers, radio = ajustar_grilla(body_gray, bands)
    scores = calcular_puntajes(body_gray, y_centers_por_banda, band_x_centers, radio)

    # baseline/peak POR BANDA, no globales: una foto de celular casi siempre tiene
    # iluminación despareja de un lado a otro de la hoja (sombra de la mano, ángulo
    # de la luz), así que "qué tan oscuro es el papel en blanco" puede variar de una
    # columna a otra. Calcularlo por banda evita que una columna más oscura por
    # iluminación (no por tinta) se lea sistemáticamente distinto a las demás.
    n_filas = len(y_centers_por_banda[0])
    resultados = []
    for bi in range(len(band_x_centers)):
        banda_scores = scores[bi * n_filas:(bi + 1) * n_filas]
        banda_vals = np.array([v for s in banda_scores for v in s.values()])
        baseline = float(np.percentile(banda_vals, 15))
        peak = float(np.percentile(banda_vals, 97))
        resultados.extend(clasificar_pregunta(s, baseline, peak) for s in banda_scores)

    return {
        "body_bgr": body_bgr,
        "y_centers": y_centers_por_banda,  # array (n_bandas, n_filas) -- grilla independiente por banda
        "band_x_centers": band_x_centers,
        "radio": radio,
        "n_bandas": len(bands),
        "resultados": resultados,      # uno por pregunta, orden banda por banda
    }
