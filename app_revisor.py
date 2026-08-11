"""
Revisor de Hojas de Respuestas — App Streamlit
- Pauta dinámica (N preguntas configurables)
- Carga de fotos desde celular o PC
- Edición de datos del alumno y corrección de respuestas dudosas
- Exporta Excel con resumen, detalle y estadísticas
"""

import io
import json
import re
import base64
import hashlib
from collections import Counter
import qrcode
import pandas as pd
import streamlit as st
import anthropic
from PIL import Image, ImageEnhance, ImageOps
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Motor OMR (opcional): si opencv/numpy no están disponibles en el entorno de
# despliegue, la app sigue funcionando 100% igual que antes con el flujo solo-IA
# — el modo OMR simplemente no aparece como opción en el sidebar.
try:
    import numpy as np
    import cv2
    import omr
    OMR_DISPONIBLE = True
except Exception:
    OMR_DISPONIBLE = False

st.set_page_config(
    page_title="Revisor de Pruebas",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="expanded",
)

VERSION_APP = "2.1.0"
FECHA_ACTUALIZACION = "2026-08-11"
DESARROLLADO_POR = "Matías Rifo V."

st.markdown("""
<style>
.corr-badge {
    display:inline-block; padding:3px 8px; border-radius:6px;
    font-size:12px; font-weight:600; margin:2px;
}
.badge-ok   { background:#dcfce7; color:#166534; }
.badge-err  { background:#fee2e2; color:#991b1b; }
.badge-dud  { background:#fef9c3; color:#92400e; }
.badge-null { background:#f3f4f6; color:#6b7280; }
div[data-testid="stDataEditor"] { font-size:13px; }
div[data-testid="stFileUploader"] > div { min-height: 120px; }
</style>
""", unsafe_allow_html=True)

# ─── Estado de sesión ────────────────────────────────────────────────
def df_pauta_vacio(n: int) -> pd.DataFrame:
    return pd.DataFrame({
        "N°": [f"P{i}" for i in range(1, n + 1)],
        "Respuesta": pd.array([None] * n, dtype="object"),
    })

def safe_key(s: str) -> str:
    """Convierte nombre de archivo en clave segura para widgets."""
    return re.sub(r'[^a-zA-Z0-9]', '_', s)

for k, v in {
    "n_preguntas": 80,
    "resultados": {},
    "correcciones": {},      # {arch: {str(num_p): letra}}
    "info_edits": {},        # {arch: {campo: valor}}
    "pauta": [],
    "fotos_pendientes": {},  # {id_unico: {nombre, bytes, mime, hash}} — subidas sin procesar aún
    "pauta_df": df_pauta_vacio(80),
    "modo_captura": "completa",  # "completa" | "solo_respuestas"
    "usar_omr": True,             # motor OMR es el método principal; la IA solo apoya en dudas e identificación
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

OPCIONES = ["A", "B", "C", "D", "E", "—"]


def prompt_dinamico(n: int, num_imagenes: int = 1, solo_respuestas: bool = False) -> str:
    num_columnas = -(-n // 20)  # división hacia arriba: esta plantilla usa columnas de 20 preguntas
    rangos = []
    for c in range(num_columnas):
        ini = c * 20 + 1
        fin = min((c + 1) * 20, n)
        rangos.append(f"columna {c+1} = preguntas {ini} a {fin}")
    descripcion_columnas = "; ".join(rangos)

    if solo_respuestas and num_imagenes >= 2:
        nota_imagenes = """IMPORTANTE — te adjunto 2 fotos que son acercamientos en zoom de la MISMA hoja, ya
recortada de antemano para mostrar ÚNICAMENTE el bloque RESPUESTAS (sin cabecera ni datos del estudiante):
1. Acercamiento de la MITAD IZQUIERDA del bloque RESPUESTAS.
2. Acercamiento de la MITAD DERECHA del bloque RESPUESTAS (con algo de superposición con la anterior).

No hay ninguna otra imagen de referencia: estas 2 fotos son todo lo que tienes de esta hoja. No busques ni
inventes nombre, cédula o folleto — no aparecen en estas fotos y tampoco se piden en la respuesta.

"""
    elif solo_respuestas:
        nota_imagenes = """IMPORTANTE — la foto adjunta ya viene recortada de antemano para mostrar ÚNICAMENTE
el bloque RESPUESTAS (sin cabecera ni datos del estudiante). No busques ni inventes nombre, cédula o folleto —
no aparecen en esta foto y tampoco se piden en la respuesta.

"""
    elif num_imagenes >= 3:
        nota_imagenes = """IMPORTANTE — te adjunto 3 fotos de la MISMA hoja de respuestas, en este orden:
1. La hoja completa (úsala para identificar al estudiante: apellidos, nombres, cédula, folleto).
2. Un acercamiento en zoom de la MITAD IZQUIERDA de esa misma hoja.
3. Un acercamiento en zoom de la MITAD DERECHA de esa misma hoja (con algo de superposición con la anterior).

Para leer el bloque RESPUESTAS, usa SIEMPRE los acercamientos (imágenes 2 y 3), NUNCA la foto completa — en
los acercamientos cada burbuja se ve mucho más grande y es mucho más fácil distinguir cuál está marcada. La
foto completa (imagen 1) es solo de referencia general y para los datos de identificación del estudiante.

"""
    else:
        nota_imagenes = ""

    if solo_respuestas:
        bloque_estructura = f"""Estás viendo SOLO el bloque "RESPUESTAS" de una hoja de respuestas de alternativas
(A/B/C/D/E) de estudiantes chilenos, recortado de la plantilla fija "PLANTILLA DE HOJA DE RESPUESTAS". Esa
plantilla organiza siempre las respuestas en columnas de 20 preguntas cada una, puestas una al lado de la otra
de izquierda a derecha: {descripcion_columnas}. Dentro de cada columna, cada fila tiene el número de pregunta
impreso a la izquierda seguido de 5 burbujas A-E."""
    else:
        bloque_estructura = f"""Eres un experto en leer hojas de respuestas de alternativas (A/B/C/D/E) de estudiantes chilenos.
Todas las hojas que vas a revisar usan siempre la misma plantilla fija "PLANTILLA DE HOJA DE RESPUESTAS", con
esta estructura exacta:

- Arriba a la izquierda, "IDENTIFICACIÓN DEL ESTUDIANTE": Apellido Paterno, Apellido Materno y Nombres,
  escritos A MANO, una letra por casillero (no son burbujas).
- Arriba a la derecha, "CÉDULA DE IDENTIDAD": los dígitos pueden venir escritos a mano en la fila superior de
  esa sección y/o marcados con burbujas (una columna de burbujas 0-9 por cada dígito). Si el texto a mano es
  legible, úsalo como fuente principal del RUT; si no, complétalo leyendo qué burbuja está marcada en cada
  columna de dígito. Ignora el recuadro "PASAPORTE" (normalmente vacío).
- Más abajo a la izquierda, casilleros separados de "N° DE FOLLETO", "SEDE", "LOCAL" y "SALA" — a menudo
  vienen vacíos; si no hay nada escrito ahí, deja "nro_folleto" como cadena vacía.
- El bloque grande "RESPUESTAS" está dividido en columnas de 20 preguntas cada una, puestas una al lado de la
  otra de izquierda a derecha: {descripcion_columnas}. Dentro de cada columna, cada fila tiene el número de
  pregunta impreso a la izquierda seguido de 5 burbujas A-E."""

    if solo_respuestas:
        schema = """{
  "respuestas": ["A","B",...],
  "dudosas": [3, 15]
}"""
        reglas_formato = f"""Reglas de formato:
- "respuestas": exactamente {n} elementos, en el mismo orden P1..P{n}. Usa null solo si la pregunta está
  realmente omitida (ninguna marca visible), no como comodín para lo que no estés seguro.
- "dudosas": números de pregunta (1 a {n}) con riesgo real de error, según el criterio simple de arriba.
- Solo el JSON, sin explicación ni comentarios adicionales."""
    else:
        schema = """{
  "apellido_paterno": "...",
  "apellido_materno": "...",
  "nombres": "...",
  "cedula": "...",
  "nro_folleto": "...",
  "respuestas": ["A","B",...],
  "dudosas": [3, 15]
}"""
        reglas_formato = f"""Reglas de formato:
- "respuestas": exactamente {n} elementos, en el mismo orden P1..P{n}. Usa null solo si la pregunta está
  realmente omitida (ninguna marca visible), no como comodín para lo que no estés seguro.
- "dudosas": números de pregunta (1 a {n}) con riesgo real de error, según el criterio simple de arriba.
- Campos de texto ilegibles o no visibles en la hoja: cadena vacía "".
- Solo el JSON, sin explicación ni comentarios adicionales."""

    return f"""{nota_imagenes}{bloque_estructura}

CÓMO ESTÁ MARCADA ESTA HOJA:
Cada pregunta tiene 5 círculos impresos que dicen A, B, C, D, E. El estudiante marca su respuesta rellenando o
rayando con lápiz mina el círculo elegido. El círculo marcado se ve MÁS OSCURO, GRIS o RELLENO comparado con
los otros 4 de la misma fila, que quedan vacíos/blancos con solo la letra impresa adentro.

REGLA DE ORO para leer cada fila: para CADA fila de 5 burbujas, compara las 5 ENTRE SÍ y elige la que se vea
más oscura, gris o rellena respecto a sus vecinas de esa misma fila — no evalúes cada burbuja aislada ni le
pidas que esté "perfectamente" rellena. Incluso en una imagen borrosa o de baja resolución, el círculo marcado
casi siempre tiene visiblemente más tono gris que los vacíos de su misma fila; esa diferencia relativa dentro
de la fila es más confiable que juzgar una sola burbuja por sí sola.

PASO 1 — Usa el número IMPRESO de cada fila, nunca el orden espacial en que la vas mirando. Por ejemplo, la
fila que dice "21" al lado de las burbujas es la pregunta 21 sin importar en qué columna esté ni si la miraste
antes o después que la fila "5" — no asumas que "la siguiente fila hacia abajo" continúa la numeración de la
columna anterior; cada columna empieza y termina en el rango indicado arriba.

PASO 2 — Registra cada respuesta en el índice del arreglo que corresponde exactamente a su número de pregunta
impreso (posición 1 = pregunta "1", posición 21 = pregunta "21", etc.), recorriendo columna por columna.

PASO 3 — Antes de responder, verifica que el arreglo "respuestas" tenga exactamente {n} elementos, uno por
cada número de pregunta impreso del 1 al {n} sin saltos ni desplazamientos. Luego vuelve a mirar por segunda
vez, aplicando otra vez la REGLA DE ORO, SOLO las preguntas donde no quedaste 100% seguro de cuál opción
marcó el estudiante, y confírmalas con calma.

CUÁNDO usar null y marcar como "dudosa" (mantén esta lista lo más corta posible): únicamente cuando, al
comparar las 5 burbujas de una fila entre sí, las 5 se ven igual de vacías (pregunta omitida) o dos se ven
igual de oscuras/rellenas entre sí (ambigüedad real, no se puede distinguir cuál de las dos es la marca). Si
una burbuja es claramente la más oscura de su fila aunque sea levemente, ESA es la respuesta — no es dudosa,
y no la dejes en null. Nunca elijas una letra al azar "porque alguna debe ser": si tras aplicar la REGLA DE
ORO dos veces sigues sin poder distinguir cuál de las 5 está marcada, usa null y agrégala a "dudosas" en vez
de adivinar — adivinar en silencio es peor que un dudoso, porque se ve como una respuesta correcta pero puede
ser falsa.

Responde ÚNICAMENTE con un JSON válido, sin texto adicional ni markdown, con esta forma exacta:

{schema}

{reglas_formato}"""


def prompt_identificacion() -> str:
    """Prompt acotado para el motor OMR: la IA NO lee burbujas de respuestas aquí
    (eso ya lo resolvió el OMR), solo transcribe los datos manuscritos/marcados
    de la cabecera de la hoja — más barato y sin riesgo de confundir tareas."""
    return """Estás viendo la cabecera de una "PLANTILLA DE HOJA DE RESPUESTAS" chilena (no el
bloque de respuestas). Contiene:
- "IDENTIFICACIÓN DEL ESTUDIANTE": Apellido Paterno, Apellido Materno y Nombres, escritos A
  MANO, una letra por casillero.
- "CÉDULA DE IDENTIDAD": dígitos que pueden venir escritos a mano y/o marcados con burbujas
  (una columna de burbujas 0-9 por cada dígito). Usa el texto a mano si es legible; si no,
  completa leyendo qué burbuja está marcada en cada columna. Ignora "PASAPORTE" si está vacío.
- Casilleros de "N° DE FOLLETO", "SEDE", "LOCAL" y "SALA" — a menudo vacíos.

Responde ÚNICAMENTE un JSON válido, sin texto adicional:
{"apellido_paterno":"...","apellido_materno":"...","nombres":"...","cedula":"...","nro_folleto":"..."}

Campos ilegibles o no visibles: cadena vacía "". Solo el JSON."""


def prompt_revision_ambiguas(preguntas: list) -> str:
    """
    Prompt para la segunda revisión focalizada: NO es el prompt completo de 80
    preguntas, es deliberadamente acotado a "cuál burbuja está marcada en este
    recorte puntual" — el motor OMR ya resolvió el resto de la hoja y solo pide
    ayuda en las filas donde no pudo decidir con confianza (marca débil, dos
    alternativas muy parecidas, o mancha que se corre hacia una burbuja vecina).
    """
    lista = ", ".join(f"P{p}" for p in preguntas)
    return f"""Te adjunto {len(preguntas)} recortes de una hoja de respuestas de alternativas
(A/B/C/D/E). Cada recorte muestra UNA fila: su número de pregunta impreso a la izquierda y
sus 5 burbujas A-E. Te los muestro en este orden: {lista}.

Un sistema automático de reconocimiento óptico ya leyó el resto de la hoja y solo tiene dudas
en estas filas puntuales (marca débil, dos alternativas con relleno parecido, o una mancha de
lápiz que se corre hacia la burbuja vecina). Tu única tarea es decidir, para cada recorte,
cuál de las 5 burbujas tiene la marca de lápiz — NO intentes resolver ni evaluar la pregunta.

Para cada recorte, compara las 5 burbujas ENTRE SÍ y elige la que se vea más oscura/rellena
respecto a sus vecinas de esa misma fila. Si de verdad las 5 se ven vacías, responde null. Si
dos burbujas están igual de marcadas y es imposible distinguir cuál es la real, responde null
también — no adivines "porque alguna debe ser".

Responde ÚNICAMENTE un JSON válido con esta forma exacta, un elemento por recorte en el mismo
orden en que te los mostré:
{{"resultados": [{{"pregunta": N, "letra": "A"}}, {{"pregunta": N, "letra": null}}, ...]}}"""


# ─── Funciones de datos ──────────────────────────────────────────────

def api_key_activa() -> str:
    try:
        return st.secrets["ANTHROPIC_API_KEY"]
    except Exception:
        return st.session_state.get("api_key_input", "")

def tiene_secret() -> bool:
    try:
        _ = st.secrets["ANTHROPIC_API_KEY"]
        return True
    except Exception:
        return False

TIPOS_MIME = {"image/jpeg":"image/jpeg","image/jpg":"image/jpeg",
              "image/png":"image/png","image/webp":"image/webp","image/heic":"image/jpeg"}

def _img_a_b64_jpeg(img: Image.Image, lado_max: int, calidad: int) -> str:
    img = img.convert("RGB")
    if max(img.size) > lado_max:
        escala = lado_max / max(img.size)
        img = img.resize((max(1,int(img.width*escala)), max(1,int(img.height*escala))), Image.Resampling.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=calidad, optimize=True)
    return base64.standard_b64encode(buf.getvalue()).decode()

def mejorar_contraste_burbujas(img: Image.Image) -> Image.Image:
    """
    Solo para los acercamientos del bloque RESPUESTAS: pasar a escala de grises quita
    ruido de color de la foto (sombras, tono del papel) y subir contraste + nitidez
    separa más el gris del rayado a lápiz del blanco de la burbuja vacía. No se aplica
    a la foto completa porque ahí conviene mantener el color para leer bien nombre/RUT.
    """
    img_gris = img.convert("L").convert("RGB")
    img_contraste = ImageEnhance.Contrast(img_gris).enhance(1.8)
    return ImageEnhance.Sharpness(img_contraste).enhance(2.0)

def abrir_imagen_corregida(datos_bytes: bytes):
    """
    Abre la imagen a su resolución ORIGINAL (sin redimensionar ni comprimir) y
    corrige la orientación EXIF: las fotos de celular suelen traer el tag EXIF
    "Orientation" en vez de venir realmente rotadas en los píxeles, y sin
    corregir esto la imagen puede llegar "acostada" 90°/180° aunque se vea
    derecha en el celular, desalineando cualquier recorte/detección posterior.
    Esto NO cubre fotos que ya vienen genuinamente rotadas en los píxeles (sin
    metadato EXIF, p.ej. tras pasar por WhatsApp o un editor que lo eliminó) —
    para esas, la corrección tiene que hacerla quien toma/recorta la foto.
    Devuelve None si la imagen no se puede decodificar.
    """
    try:
        img = Image.open(io.BytesIO(datos_bytes))
        img = ImageOps.exif_transpose(img)
        return img.convert("RGB")
    except Exception:
        return None


def preparar_imagenes(datos_bytes: bytes, mime: str, solo_respuestas: bool = False):
    """
    Claude redimensiona internamente cualquier imagen a un techo fijo de resolución
    antes de analizarla. Si las 80 preguntas van en una sola foto, ese techo se
    reparte entre las 80 y cada burbuja queda en unos pocos píxeles — insuficiente
    para distinguir con fiabilidad el rayado a lápiz. Por eso se generan 2-3 versiones
    de la misma foto: mitad izquierda / mitad derecha del bloque RESPUESTAS (siempre),
    más la foto completa para identificación del alumno (solo si el modo NO es
    "solo_respuestas") — cada mitad le da a su franja del bloque RESPUESTAS su propio
    techo de resolución completo en vez de compartirlo con el resto de la hoja.
    """
    img = abrir_imagen_corregida(datos_bytes)
    if img is None:
        return [(base64.standard_b64encode(datos_bytes).decode(), mime)]
    w, h = img.size
    margen = 0.10  # superposición horizontal para no cortar una columna justo por la mitad
    mitad = w // 2
    overlap = int(w * margen)

    if solo_respuestas:
        # La foto ya viene recortada por quien la sube para mostrar solo el bloque
        # RESPUESTAS: no hay cabecera que saltar (top=0, se aprovecha el 100% del alto)
        # ni necesidad de una "completa" (la identificación es 100% manual en este modo),
        # así que todo el presupuesto de resolución de Claude se dedica a las burbujas.
        izq = mejorar_contraste_burbujas(img.crop((0, 0, min(w, mitad + overlap), h)))
        der = mejorar_contraste_burbujas(img.crop((max(0, mitad - overlap), 0, w, h)))
        return [
            (_img_a_b64_jpeg(izq, 1568, 92), "image/jpeg"),
            (_img_a_b64_jpeg(der, 1568, 92), "image/jpeg"),
        ]

    completa = _img_a_b64_jpeg(img, 1568, 90)
    # Recorte vertical conservador: en esta plantilla la identificación del alumno
    # siempre ocupa bastante más del 10% superior de la hoja, así que descartar solo
    # ese 10% en los acercamientos no arriesga cortar filas de RESPUESTAS aunque la
    # foto venga encuadrada de forma distinta cada vez.
    top = int(h * 0.10)
    izq = mejorar_contraste_burbujas(img.crop((0, top, min(w, mitad + overlap), h)))
    der = mejorar_contraste_burbujas(img.crop((max(0, mitad - overlap), top, w, h)))
    return [
        (completa, "image/jpeg"),
        (_img_a_b64_jpeg(izq, 1568, 92), "image/jpeg"),
        (_img_a_b64_jpeg(der, 1568, 92), "image/jpeg"),
    ]

def evaluar_sospecha(res: dict) -> None:
    """
    Red de seguridad estadística: si el modelo se desalinea de fila/columna (p.ej. por
    una foto rotada o muy inclinada) tiende a leer una franja completa como si fuera
    siempre la misma burbuja, produciendo una letra que se repite muchísimo más de lo
    que un patrón real de respuestas de examen permitiría. Esto no reemplaza la
    revisión humana, pero evita que un resultado así de inverosímil pase inadvertido
    como si fuera una lectura normal.
    """
    respuestas = res.get("respuestas", [])
    no_nulas = [r.upper() for r in respuestas if r]
    n_resp = len(no_nulas)
    if n_resp == 0:
        res["sospechoso"] = True
        res["motivo_sospecha"] = "No se detectó ninguna respuesta marcada en toda la hoja."
        return
    letra_top, freq_top = Counter(no_nulas).most_common(1)[0]
    proporcion = freq_top / n_resp
    if n_resp >= 15 and proporcion >= 0.55:
        res["sospechoso"] = True
        res["motivo_sospecha"] = (
            f"{round(proporcion*100)}% de las respuestas detectadas son '{letra_top}' "
            f"({freq_top} de {n_resp}) — patrón inusual para un examen real, probable "
            f"desalineación de fila/columna. Revisa esta hoja con la foto original.")
    else:
        res["sospechoso"] = False
        res["motivo_sospecha"] = ""

def _llamar_claude(cliente, bloques_imagen: list, n: int, num_imagenes: int,
                    solo_respuestas: bool = False, texto_extra: str = "") -> dict:
    msg = cliente.messages.create(
        model="claude-sonnet-5", max_tokens=6144,
        messages=[{"role":"user","content": bloques_imagen + [
            {"type":"text","text":prompt_dinamico(n, num_imagenes, solo_respuestas) + texto_extra},
        ]}],
    )
    if msg.stop_reason == "max_tokens":
        raise ValueError("La respuesta de la IA se cortó por límite de tokens antes de terminar el JSON.")
    texto = next((b.text for b in msg.content if b.type == "text"), None)
    if texto is None:
        raise ValueError("La respuesta de la IA no incluyó texto (solo bloques de razonamiento u otro tipo).")
    texto = texto.strip()
    m = re.search(r'\{[\s\S]*\}', texto)
    if m:
        texto = m.group(0)
    res = json.loads(texto)
    resp = res.get("respuestas", [])
    res["respuestas"] = (resp + [None]*n)[:n]
    res["dudosas"] = sorted(set(res.get("dudosas", [])))
    return res

REFUERZO_REINTENTO = (
    "\n\nATENCIÓN: en un intento anterior de leer esta misma hoja, el patrón de respuestas "
    "resultó estadísticamente inverosímil para un examen real (una misma letra se repitió "
    "muchísimo más de lo esperable), lo que sugiere que te desalineaste de fila o de columna "
    "en algún punto. Vuelve a examinar la hoja desde cero, con máximo cuidado, aplicando la "
    "REGLA DE ORO fila por fila y verificando el número IMPRESO de cada fila antes de registrar "
    "su respuesta.")

def procesar_imagen(cliente, nombre: str, datos_bytes: bytes, mime: str, n: int,
                     solo_respuestas: bool = False) -> dict:
    imagenes = preparar_imagenes(datos_bytes, mime, solo_respuestas)
    bloques_imagen = [
        {"type":"image","source":{"type":"base64","media_type":mt,"data":data}}
        for data, mt in imagenes
    ]
    res = _llamar_claude(cliente, bloques_imagen, n, len(imagenes), solo_respuestas)
    evaluar_sospecha(res)
    intentos = 1
    # Un solo reintento automático cuando el primer resultado se ve estadísticamente
    # inverosímil: es barato (una llamada más) frente al costo de que el profesor
    # confíe en un resultado erróneo sin darse cuenta.
    if res["sospechoso"]:
        res2 = _llamar_claude(cliente, bloques_imagen, n, len(imagenes), solo_respuestas, REFUERZO_REINTENTO)
        evaluar_sospecha(res2)
        intentos = 2
        if not res2["sospechoso"]:
            res = res2
        elif res2.get("respuestas") != res.get("respuestas"):
            # Ambos intentos siguen sospechosos: nos quedamos con el segundo (ya vio la
            # advertencia) pero el flag queda encendido para que quede marcada a mano.
            res = res2
    res["archivo"] = nombre
    res["n_preguntas"] = n
    res["intentos"] = intentos
    res["solo_respuestas"] = solo_respuestas
    return res


# ─── Motor OMR + IA solo para dudas (beta) ────────────────────────────
# Pipeline paralelo al 100%-IA de arriba (que se deja intacto como respaldo
# garantizado): en vez de pedirle a Claude que lea las 400 burbujas de la hoja,
# un motor de visión clásica (ver omr.py) mide directamente cuánto más oscura
# está cada burbuja que el papel en blanco, comparando siempre dentro de la
# misma fila. Solo las preguntas donde ese análisis no llega a una conclusión
# confiable (marca débil, dos alternativas parecidas, mancha corrida) se
# recortan y se mandan a Claude para una segunda mirada puntual — el resto se
# resuelve sin llamar a la API. Si CUALQUIER etapa de este pipeline falla
# (imagen no decodificable, tabla no localizable, etc.), se cae automáticamente
# al flujo 100%-IA de siempre.

def _bgr_a_jpeg_b64(img_bgr, calidad: int = 95, lado_max: int = None):
    if lado_max and max(img_bgr.shape[:2]) > lado_max:
        escala = lado_max / max(img_bgr.shape[:2])
        img_bgr = cv2.resize(img_bgr, (max(1, int(img_bgr.shape[1]*escala)), max(1, int(img_bgr.shape[0]*escala))),
                              interpolation=cv2.INTER_AREA)
    ok, buf = cv2.imencode(".jpg", img_bgr, [cv2.IMWRITE_JPEG_QUALITY, calidad])
    if not ok:
        raise omr.OMRError("No se pudo codificar un recorte OMR como JPEG.")
    return base64.standard_b64encode(buf.tobytes()).decode()


def llamar_claude_identificacion(cliente, header_bgr) -> dict:
    """Segunda llamada, pequeña y acotada: solo transcribe la cabecera (nombre/RUT/folleto),
    nunca lee burbujas de respuestas — eso ya lo resolvió el motor OMR."""
    data = _bgr_a_jpeg_b64(header_bgr, calidad=90, lado_max=1024)
    msg = cliente.messages.create(
        model="claude-sonnet-5", max_tokens=512,
        messages=[{"role": "user", "content": [
            {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": data}},
            {"type": "text", "text": prompt_identificacion()},
        ]}],
    )
    texto = next((b.text for b in msg.content if b.type == "text"), None)
    if texto is None:
        return {}
    m = re.search(r'\{[\s\S]*\}', texto.strip())
    try:
        return json.loads(m.group(0) if m else texto)
    except Exception:
        return {}


def llamar_claude_revision_ambiguas(cliente, crops: list) -> dict:
    """crops: lista de (numero_pregunta, imagen_bgr). Devuelve {numero_pregunta: letra_o_None}.
    Preguntas que la IA no devuelva explícitamente quedan como None (no se adivina)."""
    if not crops:
        return {}
    preguntas = [q for q, _ in crops]
    bloques = [
        {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg",
                                      "data": _bgr_a_jpeg_b64(c, calidad=95)}}
        for _, c in crops
    ]
    msg = cliente.messages.create(
        model="claude-sonnet-5", max_tokens=1024,
        messages=[{"role": "user", "content": bloques + [
            {"type": "text", "text": prompt_revision_ambiguas(preguntas)},
        ]}],
    )
    texto = next((b.text for b in msg.content if b.type == "text"), None)
    resultado = {q: None for q in preguntas}
    if texto is None:
        return resultado
    m = re.search(r'\{[\s\S]*\}', texto.strip())
    try:
        data = json.loads(m.group(0) if m else texto)
        for item in data.get("resultados", []):
            q = item.get("pregunta")
            letra = item.get("letra")
            if q in resultado:
                resultado[q] = letra.upper() if isinstance(letra, str) and letra.upper() in LETRAS_VALIDAS else None
    except Exception:
        pass
    return resultado


def analizar_hoja_omr(datos_bytes: bytes, solo_respuestas: bool, n: int) -> dict:
    """
    Corre el motor OMR puro (sin llamadas a la API) sobre la imagen a resolución
    original. Devuelve la salida cruda de omr.analizar_imagen truncada/rellenada
    a n preguntas, más la imagen (BGR, sin comprimir) por si hace falta recortar
    preguntas puntuales para la IA. Lanza omr.OMRError si no logra ubicar una
    grilla legible (la llamante debe caer al flujo 100%-IA en ese caso).
    """
    img_pil = abrir_imagen_corregida(datos_bytes)
    if img_pil is None:
        raise omr.OMRError("La imagen no se pudo decodificar.")
    img_bgr = cv2.cvtColor(np.array(img_pil), cv2.COLOR_RGB2BGR)
    salida = omr.analizar_imagen(img_bgr, es_recorte=solo_respuestas)
    resultados = salida["resultados"]
    n_disponibles = len(resultados)
    if n_disponibles < n:
        # La tabla detectada cubre menos columnas de las que hacen falta para
        # n preguntas (p.ej. la foto no llegó a mostrar la 4ª columna, o la
        # detección de bandas no logró separarlas todas). Rellenar con
        # resultados "fantasma" es peligroso: no tienen banda/fila real en la
        # grilla, así que cualquier recorte o anotación posterior sobre esas
        # posiciones revienta con un índice fuera de rango. Es más seguro
        # tratar esto como un fallo del pipeline OMR y caer al flujo 100%-IA
        # para esta hoja en particular.
        raise omr.OMRError(
            f"La tabla detectada solo cubre {n_disponibles} de las {n} preguntas configuradas "
            "(probablemente la foto no muestra todas las columnas de RESPUESTAS).")
    salida["resultados"] = resultados[:n]
    salida["img_bgr_original"] = img_bgr
    salida["quad_respuestas"] = None if solo_respuestas else omr.detectar_bloque_respuestas(img_bgr)
    return salida


def procesar_imagen_hibrido(cliente, nombre: str, datos_bytes: bytes, mime: str, n: int,
                             solo_respuestas: bool = False) -> dict:
    try:
        salida = analizar_hoja_omr(datos_bytes, solo_respuestas, n)
    except Exception as e:
        # Cualquier fallo del pipeline OMR (tabla no localizable, imagen ilegible,
        # etc.) cae directo al flujo 100%-IA de siempre: nunca se pierde una hoja
        # por un problema del motor OMR.
        res = procesar_imagen(cliente, nombre, datos_bytes, mime, n, solo_respuestas)
        res["omr_meta"] = {"usado": False, "motivo_fallback": str(e)}
        return res

    resultados = salida["resultados"]
    body_bgr = salida["body_bgr"]
    y_centers, band_x_centers, radio = salida["y_centers"], salida["band_x_centers"], salida["radio"]
    umbral_blanco = omr.THRESHOLDS["BLANK_REVIEW_MARGIN"]

    def necesita_ia(r):
        if r["status"] in ("ambiguo", "doble_marca", "confianza_media"):
            return True
        if r["status"] == "sin_marca" and r["omr_confidence"] > umbral_blanco:
            return True
        return False

    indices_ia = [i for i, r in enumerate(resultados) if necesita_ia(r)]
    crops_ia = [(i + 1, omr.recortar_pregunta(body_bgr, y_centers, band_x_centers, radio, i, salida["n_bandas"]))
                for i in indices_ia]

    ia_resultados = {}
    if crops_ia:
        try:
            ia_resultados = llamar_claude_revision_ambiguas(cliente, crops_ia)
        except Exception:
            ia_resultados = {}  # la IA de revisión falló: esas preguntas quedan como dudosas para corrección manual

    respuestas, dudosas, metodo_por_pregunta, confianza_por_pregunta = [], [], [], []
    for i, r in enumerate(resultados):
        q = i + 1
        estado_final = r["status"]
        letra_final = r["letra"]
        metodo = "omr"
        if i in indices_ia:
            letra_ia = ia_resultados.get(q)
            if r["letra"] is None:
                # el OMR no tenía candidato: se confía en la IA si dio una respuesta
                letra_final = letra_ia
                metodo = "ia" if letra_ia else "sin_resolver"
                if letra_ia is None:
                    dudosas.append(q)
            else:
                # el OMR SÍ tenía un candidato de confianza media: si la IA coincide,
                # se confirma; si discrepa, es un conflicto real y no se adivina.
                if letra_ia and letra_ia != r["letra"]:
                    letra_final = None
                    metodo = "conflicto"
                    dudosas.append(q)
                elif letra_ia == r["letra"]:
                    metodo = "omr_confirmado_ia"
                else:  # la IA no logró determinar nada tampoco
                    metodo = "omr_no_confirmado"
                    dudosas.append(q)
            estado_final = metodo
        elif estado_final in ("ambiguo", "doble_marca"):
            dudosas.append(q)  # por si necesita_ia() no las capturó (crops_ia vacío por error de llamada)
        respuestas.append(letra_final)
        metodo_por_pregunta.append(metodo)
        confianza_por_pregunta.append(r.get("omr_confidence", 0.0))

    res = {
        "respuestas": respuestas,
        "dudosas": sorted(set(dudosas)),
        "archivo": nombre,
        "n_preguntas": n,
        "solo_respuestas": solo_respuestas,
        "apellido_paterno": "", "apellido_materno": "", "nombres": "", "cedula": "", "nro_folleto": "",
        "omr_meta": {
            "usado": True,
            "n_bandas": salida["n_bandas"],
            "metodo_por_pregunta": metodo_por_pregunta,
            "confianza_por_pregunta": confianza_por_pregunta,
            "n_omr_directo": sum(1 for m in metodo_por_pregunta if m in ("omr", "omr_confirmado_ia")),
            "n_enviadas_ia": len(crops_ia),
            "n_conflictos": sum(1 for m in metodo_por_pregunta if m == "conflicto"),
        },
    }

    # Datos del alumno: solo si la imagen es la hoja completa (en modo
    # solo_respuestas no hay cabecera que leer, igual que en el flujo 100%-IA).
    if not solo_respuestas:
        quad = salida.get("quad_respuestas")
        img_bgr = salida["img_bgr_original"]
        top_tabla = int(min(p[1] for p in quad)) if quad is not None else int(img_bgr.shape[0] * 0.5)
        header_bgr = img_bgr[0:max(1, top_tabla), :]
        try:
            datos_id = llamar_claude_identificacion(cliente, header_bgr)
            for campo in ("apellido_paterno", "apellido_materno", "nombres", "cedula", "nro_folleto"):
                res[campo] = datos_id.get(campo, "") or ""
        except Exception:
            pass  # el alumno queda sin identificar automáticamente; se completa a mano, igual que hoy

    # Diagnóstico visual: útil para depurar/confiar en el motor OMR sin tener que
    # revisar pregunta por pregunta.
    resultados_diag = []
    for i, r in enumerate(resultados):
        rd = dict(r)
        rd["letra"] = respuestas[i]
        rd["status"] = metodo_por_pregunta[i]
        resultados_diag.append(rd)
    diag_bgr = omr.anotar_diagnostico(body_bgr, y_centers, band_x_centers, radio, resultados_diag)
    ok, buf = cv2.imencode(".jpg", diag_bgr, [cv2.IMWRITE_JPEG_QUALITY, 85])
    res["omr_diagnostico_bytes"] = buf.tobytes() if ok else None

    evaluar_sospecha(res)
    # Último respaldo: si a pesar de OMR + revisión IA el resultado final sigue
    # viéndose estadísticamente inverosímil, se descarta todo y se usa el flujo
    # 100%-IA completo (con su propio reintento) como red de seguridad final.
    if res["sospechoso"]:
        res_ia = procesar_imagen(cliente, nombre, datos_bytes, mime, n, solo_respuestas)
        if not res_ia.get("sospechoso"):
            res_ia["omr_meta"] = {**res["omr_meta"], "usado": True, "descartado_por_sospechoso": True}
            return res_ia
    return res


def datos_efectivos(arch: str) -> dict:
    """Datos originales del alumno + ediciones manuales."""
    base = dict(st.session_state.resultados[arch])
    base.update(st.session_state.info_edits.get(arch, {}))
    return base

def respuestas_efectivas(arch: str) -> list:
    base = st.session_state.resultados[arch]["respuestas"].copy()
    for idx_str, letra in st.session_state.correcciones.get(arch, {}).items():
        idx = int(idx_str) - 1
        if 0 <= idx < len(base):
            base[idx] = None if letra == "—" else letra
    return base

def calcular(respuestas, pauta):
    co = inc = om = 0
    err = []
    for i, (r, p) in enumerate(zip(respuestas, pauta), 1):
        if not p: continue
        if r is None: om += 1
        elif r.upper() == p.upper(): co += 1
        else: inc += 1; err.append(i)
    return co, inc, om, err

LETRAS_VALIDAS = {"A", "B", "C", "D", "E"}

def pauta_desde_df(df: pd.DataFrame) -> list:
    return [
        r.strip().upper() if r and r.strip().upper() in LETRAS_VALIDAS else None
        for r in df["Respuesta"].fillna("").tolist()
    ]

def guardar_info_edit(arch, campo, valor):
    if arch not in st.session_state.info_edits:
        st.session_state.info_edits[arch] = {}
    st.session_state.info_edits[arch][campo] = valor.strip()

def guardar_correccion(arch, num_p, nueva, resp_orig):
    if nueva != (resp_orig or "—"):
        if arch not in st.session_state.correcciones:
            st.session_state.correcciones[arch] = {}
        st.session_state.correcciones[arch][str(num_p)] = nueva
    else:
        st.session_state.correcciones.get(arch, {}).pop(str(num_p), None)


# ─── Excel ───────────────────────────────────────────────────────────

def generar_excel(pauta: list, curso: str) -> bytes:
    n = len(pauta)
    wb = Workbook()
    bd = Border(left=Side(style="thin"),right=Side(style="thin"),
                top=Side(style="thin"),bottom=Side(style="thin"))
    AZ=PatternFill("solid",fgColor="2563EB"); AZD=PatternFill("solid",fgColor="1E3A5F")
    VE=PatternFill("solid",fgColor="DCFCE7"); RO=PatternFill("solid",fgColor="FECACA")
    AM=PatternFill("solid",fgColor="FEF08A"); AZC=PatternFill("solid",fgColor="DBEAFE")
    GR=PatternFill("solid",fgColor="F3F4F6")
    total_p = sum(1 for p in pauta if p)

    # Hoja 1 — Resumen (sin celdas combinadas: cada fila es autocontenida para
    # poder copiar/pegar y unificar varios Excel exportados en un archivo maestro,
    # filtrando por Curso y por alumno)
    ws1 = wb.active; ws1.title = "Resumen"
    enc = ["Curso","N°","Ap. Paterno","Ap. Materno","Nombres","Cédula","Folleto",
           "Correctas","Incorrectas","Omitidas","Puntaje %",
           "Preguntas incorrectas","Dudosas corregidas","Alerta calidad"]
    COL_PUNTAJE = 11
    COL_ALERTA = 14
    COLS_TEXTO = {1,3,4,5,6,7,14}
    fe = 1
    for c,h in enumerate(enc,1):
        cell=ws1.cell(fe,c,h); cell.font=Font(bold=True,color="FFFFFF")
        cell.fill=AZ; cell.alignment=Alignment(horizontal="center",wrap_text=True)
        cell.border=bd

    for rn,(arch,_) in enumerate(st.session_state.resultados.items(),1):
        datos = datos_efectivos(arch)
        ref   = respuestas_efectivas(arch)
        co,inc,om,err = calcular(ref, pauta[:len(ref)])
        pct = round(co/total_p*100,1) if total_p else 0
        res_orig = st.session_state.resultados[arch]
        dud = res_orig.get("dudosas",[])
        corr= st.session_state.correcciones.get(arch,{})
        cn  = len([k for k in corr if int(k) in dud])
        alerta = res_orig.get("motivo_sospecha","") if res_orig.get("sospechoso") else "—"
        vals=[curso or "—", rn,
              datos.get("apellido_paterno",""),datos.get("apellido_materno",""),
              datos.get("nombres",""),datos.get("cedula",""),datos.get("nro_folleto",""),
              co,inc,om,pct,
              ", ".join(str(e) for e in err) if err else "—",
              f"{cn} de {len(dud)}" if dud else "—",
              alerta]
        for c,v in enumerate(vals,1):
            cell=ws1.cell(fe+rn,c,v); cell.border=bd
            cell.alignment=Alignment(horizontal="left" if c in COLS_TEXTO else "center",wrap_text=True)
            if c==COL_PUNTAJE: cell.fill=VE if pct>=70 else (AM if pct>=50 else RO)
            if c==COL_ALERTA and alerta!="—": cell.fill=RO; cell.font=Font(bold=True,color="991B1B")
    for i,w in enumerate([20,4,16,16,22,14,9,10,11,10,10,38,18,44],1):
        ws1.column_dimensions[get_column_letter(i)].width=w
    ws1.row_dimensions[fe].height=22
    ws1.freeze_panes = "A2"

    # Hoja 2 — Detalle respuestas
    ws2 = wb.create_sheet("Detalle respuestas")
    cab=["Alumno"]+[f"P{i}" for i in range(1,n+1)]
    for c,h in enumerate(cab,1):
        cell=ws2.cell(1,c,h); cell.font=Font(bold=True,color="FFFFFF")
        cell.fill=AZD; cell.alignment=Alignment(horizontal="center"); cell.border=bd
    ws2.cell(2,1,"PAUTA").font=Font(bold=True)
    for i,p in enumerate(pauta,2):
        cell=ws2.cell(2,i,p or "?"); cell.fill=AZC
        cell.alignment=Alignment(horizontal="center"); cell.border=bd

    for fila,(arch,_) in enumerate(st.session_state.resultados.items(),3):
        datos=datos_efectivos(arch); ref=respuestas_efectivas(arch)
        corr=st.session_state.correcciones.get(arch,{})
        dud_set=set(st.session_state.resultados[arch].get("dudosas",[]))
        nombre=f"{datos.get('apellido_paterno','')} {datos.get('nombres','')}"
        ws2.cell(fila,1,nombre).border=bd
        for i,r in enumerate(ref,1):
            cell=ws2.cell(fila,i+1,r or "—")
            cell.alignment=Alignment(horizontal="center"); cell.border=bd
            corr_esta=str(i) in corr and i in dud_set
            if r is None: cell.fill=GR
            elif i<=len(pauta) and pauta[i-1] and r.upper()==pauta[i-1].upper():
                cell.fill=VE
                if corr_esta: cell.font=Font(bold=True,color="166534")
            else:
                cell.fill=RO
                if corr_esta: cell.font=Font(bold=True,color="991B1B")
            if i in dud_set and not corr_esta:
                cell.fill=AM; cell.font=Font(bold=True,color="92400E")
    ws2.column_dimensions["A"].width=26
    for col in range(2,n+2): ws2.column_dimensions[get_column_letter(col)].width=4.2
    ws2.row_dimensions[1].height=24

    # Hoja 3 — Estadísticas
    ws3=wb.create_sheet("Estadísticas")
    ws3["A1"]=f"Estadísticas — {curso or 'Curso'}"; ws3["A1"].font=Font(bold=True,size=14)
    ws3["A2"]=f"Total de preguntas evaluadas: {total_p}"
    todos_pct=[]
    for arch in st.session_state.resultados:
        ref=respuestas_efectivas(arch)
        co,_,_,_=calcular(ref,pauta[:len(ref)])
        todos_pct.append(round(co/total_p*100,1) if total_p else 0)
    for f,(e,v) in enumerate([
        ("Total alumnos",len(todos_pct)),
        ("Promedio (%)",round(sum(todos_pct)/len(todos_pct),1) if todos_pct else 0),
        ("Máximo (%)",max(todos_pct) if todos_pct else 0),
        ("Mínimo (%)",min(todos_pct) if todos_pct else 0),
        ("Sobre 60%",sum(1 for p in todos_pct if p>=60)),
        ("Bajo 60%",sum(1 for p in todos_pct if p<60)),
    ],4):
        ws3.cell(f,1,e).font=Font(bold=True); ws3.cell(f,2,v)
    ws3.column_dimensions["A"].width=34; ws3.column_dimensions["B"].width=14

    # Hoja 4 — Pauta
    ws4=wb.create_sheet("Pauta utilizada")
    ws4["A1"]=f"Pauta — {n} preguntas"; ws4["A1"].font=Font(bold=True,size=12)
    cols_f=5
    for idx,p in enumerate(pauta):
        fila=(idx//cols_f)+3; cb=(idx%cols_f)*2+1
        ws4.cell(fila,cb,f"P{idx+1}").font=Font(bold=True,color="1E3A5F")
        cell=ws4.cell(fila,cb+1,p or "—")
        cell.alignment=Alignment(horizontal="center")
        if p: cell.fill=AZC
    for col in range(1,cols_f*2+1): ws4.column_dimensions[get_column_letter(col)].width=7

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.read()

def generar_qr(url: str) -> bytes:
    qr=qrcode.QRCode(box_size=6,border=2); qr.add_data(url); qr.make(fit=True)
    img=qr.make_image(fill_color="black",back_color="white")
    buf=io.BytesIO(); img.save(buf,format="PNG"); buf.seek(0)
    return buf.read()


# ═══════════════════════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## ⚙️ Configuración")
    if tiene_secret():
        st.success("🔑 API Key desde secrets")
    else:
        st.session_state["api_key_input"] = st.text_input(
            "API Key de Anthropic", type="password", placeholder="sk-ant-...")
    curso = st.text_input("Nombre del curso", placeholder="Ej: 3°A — Historia 2026")
    st.markdown("---")
    st.markdown("**Número de preguntas**")
    n_prev = st.session_state["n_preguntas"]
    hay_procesadas = bool(st.session_state["resultados"])
    n_nuevo = st.number_input("Preguntas", min_value=1, max_value=120,
                               value=n_prev, step=1, label_visibility="collapsed",
                               disabled=hay_procesadas)
    if hay_procesadas:
        st.caption("🔒 Bloqueado: ya hay hojas procesadas con este N°. "
                   "Usa **Limpiar todo** antes de cambiarlo (cambiarlo a medio camino "
                   "descuadra el puntaje de las hojas ya analizadas).")
    elif n_nuevo != n_prev:
        st.session_state["n_preguntas"] = n_nuevo
        st.session_state["pauta_df"] = df_pauta_vacio(n_nuevo)
        st.session_state["pauta"] = []
        st.rerun()
    st.caption(f"Configurado para **{st.session_state['n_preguntas']} preguntas**")
    st.markdown("---")
    st.markdown("**Modo de carga de fotos**")
    opciones_modo = {
        "completa": "📄 Hoja completa (recorte automático)",
        "solo_respuestas": "✂️ Solo bloque RESPUESTAS (ya recortado por ti)",
    }
    modo_prev = st.session_state["modo_captura"]
    modo_nuevo = st.radio(
        "Modo de carga", options=list(opciones_modo.keys()),
        format_func=lambda k: opciones_modo[k],
        index=list(opciones_modo.keys()).index(modo_prev),
        label_visibility="collapsed", disabled=hay_procesadas,
    )
    if hay_procesadas:
        st.caption("🔒 Bloqueado: ya hay hojas procesadas con este modo. Usa **Limpiar todo** antes de cambiarlo.")
    elif modo_nuevo != modo_prev:
        st.session_state["modo_captura"] = modo_nuevo
        st.rerun()
    if st.session_state["modo_captura"] == "solo_respuestas":
        st.caption("✂️ Sube la foto recortada para mostrar **solo** el bloque RESPUESTAS (sin cabecera). "
                   "Nombre, RUT y folleto quedan en blanco — los completas a mano en **Revisar y corregir**. "
                   "Es el modo más preciso: toda la resolución se dedica a las burbujas.")
    else:
        st.caption("📄 Sube la foto de la hoja completa; la app recorta e identifica al alumno automáticamente.")
    st.markdown("---")
    st.markdown("**Motor de lectura**")
    if not OMR_DISPONIBLE:
        st.caption("🤖 Solo IA (motor OMR no disponible en este entorno).")
    else:
        omr_prev = st.session_state["usar_omr"]
        omr_nuevo = st.toggle(
            "🔬 Motor OMR (método principal) + IA de apoyo",
            value=omr_prev, disabled=hay_procesadas,
            help="Método principal de esta app: mide directamente qué tan oscura está cada burbuja "
                 "(visión clásica, sin llamar a la API) y SOLO recurre a Claude para las preguntas donde "
                 "no queda claro cuál está marcada, y para transcribir nombre/RUT. Más rápido, más barato "
                 "y — validado contra fotos reales — más preciso que pedirle a la IA que lea las 400 "
                 "burbujas de una hoja completa. Si en algún punto falla, esa hoja cae automáticamente al "
                 "modo 100% IA de respaldo, sin perderse.",
        )
        if hay_procesadas:
            st.caption("🔒 Bloqueado: ya hay hojas procesadas con este motor. Usa **Limpiar todo** antes de cambiarlo.")
        elif omr_nuevo != omr_prev:
            st.session_state["usar_omr"] = omr_nuevo
            st.rerun()
        if st.session_state["usar_omr"]:
            st.caption("🔬 **Configuración recomendada.** La mayoría de las preguntas se resuelven por análisis "
                       "directo de imagen (sin llamar a la API); solo las dudosas van a Claude, junto con "
                       "nombre/RUT. Puedes revisar el diagnóstico de cada hoja (qué burbuja detectó y con qué "
                       "confianza) en **Revisar y corregir**.")
        else:
            st.caption("🤖 Modo 100% IA (el que usaba la app antes del motor OMR) — más lento y más caro por "
                       "hoja. Útil para comparar o si el motor OMR falla repetidamente con tus fotos.")
    st.markdown("---")
    if st.button("🗑️ Limpiar todo", use_container_width=True):
        nn = st.session_state["n_preguntas"]
        st.session_state.update({
            "resultados":{}, "correcciones":{}, "info_edits":{},
            "fotos_pendientes":{}, "pauta":[], "pauta_df":df_pauta_vacio(nn),
        })
        st.rerun()
    st.caption(f"Versión {VERSION_APP} · Actualizado {FECHA_ACTUALIZACION}")
    st.caption(f"Desarrollado por {DESARROLLADO_POR}")


# ═══════════════════════════════════════════════════════════════════════
# CUERPO
# ═══════════════════════════════════════════════════════════════════════
n = st.session_state["n_preguntas"]
st.markdown("# 📝 Revisor de Hojas de Respuestas")

tab_pauta, tab_cargar, tab_revisar, tab_exportar = st.tabs([
    "📋 Pauta de respuestas",
    "📤 Cargar fotos",
    "✏️ Revisar y corregir",
    "📊 Exportar Excel",
])

# ══ PAUTA ════════════════════════════════════════════════════════════
with tab_pauta:
    st.markdown(f"### Respuestas correctas — {n} preguntas")
    st.caption("💡 Recomendado: usa **Importar rápido** a la derecha (pega todas las respuestas de una vez) "
               "— es más rápido y evita el parpadeo/reseteo que puede ocurrir al editar celda por celda muy "
               "rápido en la tabla. Usa la tabla solo para corregir 1 o 2 celdas puntuales.")
    col_t, col_i = st.columns([2, 1], gap="large")

    with col_i:
        st.markdown("**Importar rápido**")
        texto_bulk = st.text_area("Importar", height=200, label_visibility="collapsed",
            placeholder=f"A\nB\nC\n...\no bien: A,B,C,D,E,...\n({n} respuestas)")
        if st.button("⬇️ Cargar al grid", use_container_width=True):
            letras=[l.upper() for l in re.findall(r'[AaBbCcDdEe]', texto_bulk)]
            if letras:
                st.session_state["pauta_df"] = pd.DataFrame({
                    "N°":[f"P{i}" for i in range(1,n+1)],
                    "Respuesta": pd.array((letras+[None]*n)[:n], dtype="object"),
                })
                st.success(f"✓ {min(len(letras),n)} respuestas cargadas")
                st.rerun()
            else:
                st.error("No se encontraron letras A–E válidas.")
        st.markdown("---")
        pauta_actual = pauta_desde_df(st.session_state["pauta_df"])
        completadas  = sum(1 for p in pauta_actual if p)
        st.metric("Con respuesta", f"{completadas} / {n}")
        if completadas == n:
            st.success("✅ Pauta completa")
        elif completadas > 0:
            falt=[i+1 for i,p in enumerate(pauta_actual) if not p]
            st.warning(f"Faltan {n-completadas}: {', '.join(f'P{x}' for x in falt[:5])}{'...' if len(falt)>5 else ''}")

    with col_t:
        df_ed = st.data_editor(
            st.session_state["pauta_df"],
            column_config={
                "N°": st.column_config.TextColumn("N°", disabled=True, width="small"),
                "Respuesta": st.column_config.SelectboxColumn(
                    "Respuesta", options=["A","B","C","D","E"], width="small", required=False),
            },
            hide_index=True, height=min(38*n+40, 560),
            use_container_width=True, key="editor_pauta", num_rows="fixed",
        )
        st.session_state["pauta_df"] = df_ed
        st.session_state["pauta"]    = pauta_desde_df(df_ed)

    if any(st.session_state["pauta"]):
        with st.expander("👁️ Vista previa"):
            html="<div style='line-height:2.2;'>"
            for i,p in enumerate(st.session_state["pauta"],1):
                bg="#dcfce7" if p else "#f3f4f6"; fg="#166534" if p else "#9ca3af"
                html+=(f'<span style="display:inline-block;width:50px;margin:2px;'
                       f'background:{bg};color:{fg};border-radius:4px;'
                       f'padding:2px 4px;font-size:12px;font-weight:600;">'
                       f'P{i}:{p or "?"}</span>')
            html+="</div>"
            st.markdown(html, unsafe_allow_html=True)


# ══ CARGAR FOTOS ═════════════════════════════════════════════════════
with tab_cargar:
    try:
        app_url = st.secrets.get("APP_URL","")
    except Exception:
        app_url = ""

    modo_actual = st.session_state["modo_captura"]

    if app_url:
        cq, ci = st.columns([1,2], gap="large")
        with cq:
            st.markdown("### 📱 Desde el celular")
            st.image(generar_qr(app_url), width=180, caption="Escanea para abrir en tu celular")
        with ci:
            st.markdown("### Instrucciones")
            if modo_actual == "solo_respuestas":
                st.markdown("""
1. Escanea el QR o entra a la URL en el navegador del celular
2. Toca **Cargar fotos** arriba
3. Toca **Upload** → elige **Cámara** o **Galería**
4. **Recorta o encuadra la foto para que muestre SOLO el bloque RESPUESTAS**, sin la
   cabecera de nombre/RUT — que ocupe todo el encuadre, sin mesa ni fondo alrededor
5. Foto derecha (no en ángulo), buena luz, sin sombras sobre el papel
6. Sube varias antes de procesar; nombre y RUT los completas a mano después
""")
            else:
                st.markdown("""
1. Escanea el QR o entra a la URL en el navegador del celular
2. Toca **Cargar fotos** arriba
3. Toca **Upload** → elige **Cámara** o **Galería**
4. **Acerca el celular a la hoja** hasta que ocupe todo el encuadre (sin mesa, ropa
   ni pies alrededor) — con 80 preguntas cada burbuja es diminuta, así que entre
   más grande se vea la hoja en la foto, más fácil es distinguir cuál está marcada
5. **Que se vean las 4 columnas completas del bloque RESPUESTAS** — si la foto corta
   una columna, esa hoja no se puede leer por OMR y se procesa más lento (solo IA)
6. Foto derecha (no en ángulo), buena luz, sin sombras sobre el papel
7. Sube varias antes de procesar
""")
    else:
        st.info("Desde el celular: entra a la URL de esta app en el navegador. "
                "Al tocar Upload el celular ofrece abrir la cámara directamente. "
                "Agrega `APP_URL` a Secrets de Streamlit Cloud para ver un QR aquí.")

    st.markdown("---")
    st.markdown(f"### Sube las fotos  ·  *{n} preguntas por prueba*")
    st.caption("En el celular puedes tocar el recuadro varias veces para tomar una foto a la vez: "
               "cada una queda guardada aunque la cámara se abra de nuevo.")
    if modo_actual == "solo_respuestas":
        st.caption("✂️ **Clave para que el motor OMR lea bien:** la foto debe mostrar el bloque RESPUESTAS "
                   "COMPLETO (las 4 columnas, sin cabecera de nombre/RUT), encuadrada derecha, sin fondo "
                   "alrededor y con buena luz.")
    else:
        st.caption("📸 **Clave para que el motor OMR lea bien:** acerca la cámara y llena el encuadre con la "
                   "hoja completa, con las 4 columnas de RESPUESTAS visibles, foto derecha y con buena luz — "
                   "sin fondo alrededor. La IA solo entra a apoyar en las preguntas puntuales donde el motor "
                   "OMR queda con dudas, así que la calidad de la foto sigue importando.")
    st.caption("💡 Recomendado: procesa de a **15–20 fotos por lote** (no 40–80 de una vez). "
               "Así el progreso no se pierde si el celular se bloquea o hay corte de conexión, "
               "y luego puedes exportar cada lote y unirlos en un Excel maestro.")
    archivos = st.file_uploader("Fotos", type=["jpg","jpeg","png","webp"],
                                 accept_multiple_files=True, label_visibility="collapsed",
                                 key="uploader_fotos")

    # Cada foto seleccionada se guarda de inmediato en session_state (por hash de
    # contenido, no por nombre de archivo). Esto evita perder fotos anteriores cuando
    # el celular reabre la cámara y reemplaza la selección del input, y evita
    # colisiones cuando la cámara reutiliza el mismo nombre genérico (ej. "image.jpg").
    if archivos:
        hashes_conocidos = {v["hash"] for v in st.session_state.fotos_pendientes.values()}
        hashes_conocidos |= {d.get("hash") for d in st.session_state.resultados.values()}
        agregadas = 0
        for f in archivos:
            contenido = f.read()
            h = hashlib.md5(contenido).hexdigest()
            if h not in hashes_conocidos:
                # Se guarda tal cual (sin comprimir aquí): el recorte en mitades y la
                # compresión ocurren recién al procesar, a partir de la resolución
                # original, para maximizar el detalle de cada acercamiento.
                idx = len(st.session_state.fotos_pendientes) + len(st.session_state.resultados) + 1
                st.session_state.fotos_pendientes[f"foto_{idx:03d}"] = {
                    "nombre": f.name, "bytes": contenido,
                    "mime": TIPOS_MIME.get(f.type, "image/jpeg"), "hash": h,
                }
                hashes_conocidos.add(h)
                agregadas += 1
        if agregadas:
            st.rerun()

    # Mensaje del último procesamiento: se guarda en session_state porque el
    # st.rerun() de más abajo redibuja la pantalla de inmediato y borraría un
    # st.success/st.error mostrado en la misma corrida antes de que se alcance a leer.
    if st.session_state.get("mensaje_proceso"):
        tipo, texto = st.session_state.pop("mensaje_proceso")
        (st.error if tipo == "error" else st.success)(texto)

    pendientes = st.session_state.fotos_pendientes
    total_proc = len(st.session_state.resultados)

    if pendientes or total_proc:
        c1, c2 = st.columns(2)
        if total_proc: c1.success(f"✓ {total_proc} ya procesadas")
        if pendientes: c2.warning(f"⏳ {len(pendientes)} nuevas por procesar")

    if pendientes:
        with st.expander(f"📋 Fotos por procesar ({len(pendientes)})"):
            for id_unico, foto in pendientes.items():
                st.caption(f"• {foto['nombre']}")
        key = api_key_activa()
        if not key:
            st.error("Ingresa tu API Key en el panel lateral.")
        elif st.button(f"🚀 Procesar {len(pendientes)} hoja(s)", type="primary", use_container_width=True):
            cliente = anthropic.Anthropic(api_key=key, timeout=90.0, max_retries=1)
            prog = st.progress(0, text="Iniciando...")
            errs = []
            items = list(pendientes.items())
            for i, (id_unico, foto) in enumerate(items):
                prog.progress(i/len(items), text=f"Procesando {foto['nombre']} ({i+1}/{len(items)})...")
                try:
                    solo_resp = (modo_actual == "solo_respuestas")
                    if st.session_state.get("usar_omr") and OMR_DISPONIBLE:
                        res = procesar_imagen_hibrido(cliente, foto["nombre"], foto["bytes"], foto["mime"], n,
                                                       solo_respuestas=solo_resp)
                    else:
                        res = procesar_imagen(cliente, foto["nombre"], foto["bytes"], foto["mime"], n,
                                               solo_respuestas=solo_resp)
                    res["hash"] = foto["hash"]
                    st.session_state.resultados[id_unico] = res
                    st.session_state.fotos_pendientes.pop(id_unico, None)
                except anthropic.APITimeoutError:
                    errs.append(f"{foto['nombre']}: tiempo de espera agotado (conexión lenta) — quedó en la cola, reintenta.")
                except Exception as e:
                    errs.append(f"{foto['nombre']}: {e} — quedó en la cola, reintenta.")
            prog.progress(1.0, text="¡Completado!")
            if errs:
                st.session_state["mensaje_proceso"] = ("error", "Errores:\n"+"\n".join(errs))
            else:
                st.session_state["mensaje_proceso"] = ("success", f"✅ {len(items)} procesadas. Ve a **Revisar y corregir**.")
            st.rerun()
    elif total_proc:
        st.success("✅ Todas las fotos cargadas. Ve a **Revisar y corregir**.")
    else:
        st.markdown("""
        <div style="border:2px dashed #d1d5db;border-radius:16px;padding:2.5rem;
                    text-align:center;color:#6b7280;margin-top:1rem;">
            <div style="font-size:3rem;">📷</div>
            <div style="font-size:16px;margin-top:10px;font-weight:500;">
                Arrastra fotos aquí o toca para seleccionar</div>
            <div style="font-size:13px;margin-top:6px;">JPG · PNG · WEBP · Desde PC o celular</div>
        </div>""", unsafe_allow_html=True)


# ══ REVISAR Y CORREGIR ═══════════════════════════════════════════════
with tab_revisar:
    pauta = st.session_state["pauta"]

    if not st.session_state.resultados:
        st.info("Aún no hay hojas procesadas. Ve a **Cargar fotos**.")
    else:
        total_alumnos = len(st.session_state.resultados)
        pendientes = sum(
            1 for a,d in st.session_state.resultados.items()
            if [x for x in d.get("dudosas",[]) if str(x) not in st.session_state.correcciones.get(a,{})]
        )
        sospechosas = sum(1 for d in st.session_state.resultados.values() if d.get("sospechoso"))
        sin_id = sum(
            1 for a, d in st.session_state.resultados.items()
            if not (datos_efectivos(a).get("apellido_paterno") or datos_efectivos(a).get("nombres")
                     or datos_efectivos(a).get("cedula"))
        )
        m1,m2,m3,m4,m5 = st.columns(5)
        m1.metric("Alumnos procesados", total_alumnos)
        m2.metric("Con dudas pendientes", pendientes)
        m3.metric("Preguntas por prueba", n)
        m4.metric("🚨 Patrón sospechoso", sospechosas)
        m5.metric("🆔 Sin identificar", sin_id)
        if sospechosas:
            st.error(f"🚨 {sospechosas} hoja(s) con un patrón de respuestas estadísticamente inverosímil "
                     "(revisadas dos veces por la IA y aun así el resultado es raro). Ábrelas más abajo y "
                     "compáralas manualmente con la foto original antes de confiar en su puntaje.")
        if sin_id:
            st.warning(f"🆔 {sin_id} hoja(s) sin apellido, nombre ni cédula — complétalas más abajo antes de exportar.")
        st.markdown("---")

        for arch, datos_orig in st.session_state.resultados.items():
            datos     = datos_efectivos(arch)
            dudosas   = datos_orig.get("dudosas", [])
            corr_arch = st.session_state.correcciones.get(arch, {})
            pend_este = [d for d in dudosas if str(d) not in corr_arch]
            sk        = safe_key(arch)

            ref = respuestas_efectivas(arch)
            if pauta:
                co,inc,om,_ = calcular(ref, pauta[:len(ref)])
                total_p = sum(1 for p in pauta if p)
                pct = round(co/total_p*100,1) if total_p else 0
                puntaje_str = f"**{pct}%** ({co}/{total_p})"
            else:
                puntaje_str = "_(sin pauta)_"

            es_sospechoso  = bool(datos_orig.get("sospechoso"))
            sin_identificar = not (datos.get("apellido_paterno") or datos.get("nombres") or datos.get("cedula"))
            icono = "🚨" if es_sospechoso else ("🆔" if sin_identificar else ("⚠️" if pend_este else "✅"))
            nombre_titulo = (f"{datos.get('apellido_paterno','')} {datos.get('nombres','')} — {datos.get('cedula','')}"
                              if not sin_identificar
                              else f"*(sin identificar — {datos_orig.get('archivo', arch)})*")
            with st.expander(
                f"{icono} {nombre_titulo} | {puntaje_str}",
                expanded=bool(pend_este or es_sospechoso or sin_identificar)
            ):
                if es_sospechoso:
                    st.error(f"🚨 **Patrón sospechoso:** {datos_orig.get('motivo_sospecha','')}")
                if sin_identificar:
                    st.warning("🆔 Falta identificar al alumno — completa al menos apellido paterno o cédula abajo.")

                # ── Diagnóstico OMR ──────────────────────────────────
                omr_meta = datos_orig.get("omr_meta")
                if omr_meta and omr_meta.get("usado"):
                    n_total = len(ref)
                    pct_omr = round(omr_meta["n_omr_directo"] / n_total * 100) if n_total else 0
                    st.caption(
                        f"🔬 Motor OMR: **{pct_omr}%** resuelto directo por análisis de imagen "
                        f"({omr_meta['n_omr_directo']}/{n_total}) · {omr_meta['n_enviadas_ia']} enviadas a IA"
                        + (f" · ⚠️ {omr_meta['n_conflictos']} conflicto(s) OMR↔IA" if omr_meta.get('n_conflictos') else "")
                    )
                    if datos_orig.get("omr_diagnostico_bytes"):
                        with st.expander("🔬 Ver diagnóstico OMR (qué burbuja detectó y con qué confianza)"):
                            st.caption("🟢 alta confianza (OMR)  ·  🔵 revisada/confirmada por IA  ·  🔴 conflicto o sin resolver")
                            st.image(datos_orig["omr_diagnostico_bytes"], use_container_width=True)
                elif omr_meta and not omr_meta.get("usado"):
                    st.caption(f"🔬 Motor OMR no se pudo usar en esta hoja ({omr_meta.get('motivo_fallback','?')}) "
                               "— se procesó con el flujo 100% IA de siempre.")

                # ── Datos editables del alumno ──────────────────────
                if datos_orig.get("solo_respuestas"):
                    st.markdown("**Datos del alumno** *(modo solo-respuestas: complétalos a mano)*")
                else:
                    st.markdown("**Datos del alumno** *(edita si Claude leyó mal algún campo)*")
                ie = st.session_state.info_edits.get(arch, {})

                r1c1, r1c2, r1c3 = st.columns(3)
                with r1c1:
                    v = st.text_input("Apellido paterno",
                        value=ie.get("apellido_paterno", datos_orig.get("apellido_paterno","")),
                        key=f"ap_{sk}")
                    guardar_info_edit(arch, "apellido_paterno", v)
                with r1c2:
                    v = st.text_input("Apellido materno",
                        value=ie.get("apellido_materno", datos_orig.get("apellido_materno","")),
                        key=f"am_{sk}")
                    guardar_info_edit(arch, "apellido_materno", v)
                with r1c3:
                    v = st.text_input("Nombres",
                        value=ie.get("nombres", datos_orig.get("nombres","")),
                        key=f"no_{sk}")
                    guardar_info_edit(arch, "nombres", v)

                r2c1, r2c2, r2c3 = st.columns(3)
                with r2c1:
                    v = st.text_input("Cédula / RUT",
                        value=ie.get("cedula", datos_orig.get("cedula","")),
                        key=f"ce_{sk}")
                    guardar_info_edit(arch, "cedula", v)
                with r2c2:
                    v = st.text_input("N° de folleto",
                        value=ie.get("nro_folleto", datos_orig.get("nro_folleto","")),
                        key=f"fo_{sk}")
                    guardar_info_edit(arch, "nro_folleto", v)
                with r2c3:
                    st.markdown("&nbsp;")  # espaciador

                # ── Respuestas dudosas ──────────────────────────────
                if dudosas:
                    st.markdown("---")
                    st.markdown(f"**Preguntas dudosas:** {', '.join(f'P{d}' for d in dudosas)}")
                    st.caption("Selecciona la respuesta correcta para cada una:")

                    # Renderizar en grupos de 8 para evitar bug de columnas
                    CHUNK = 8
                    for chunk_start in range(0, len(dudosas), CHUNK):
                        chunk = dudosas[chunk_start:chunk_start+CHUNK]
                        cols  = st.columns(len(chunk))
                        for j, num_p in enumerate(chunk):
                            resp_orig = (datos_orig["respuestas"][num_p-1]
                                         if num_p <= len(datos_orig["respuestas"]) else None)
                            # Leer valor guardado en session_state directamente para evitar blank
                            widget_key = f"dud_{sk}_{num_p}"
                            saved_val  = corr_arch.get(str(num_p), resp_orig or "—")
                            idx_opcion = OPCIONES.index(saved_val) if saved_val in OPCIONES else 5

                            with cols[j]:
                                nueva = st.selectbox(
                                    f"P{num_p}",
                                    options=OPCIONES,
                                    index=idx_opcion,
                                    key=widget_key,
                                )
                                guardar_correccion(arch, num_p, nueva, resp_orig)
                else:
                    st.markdown("---")
                    st.success("Sin preguntas dudosas.")

                # ── Vista detalle de respuestas ─────────────────────
                if pauta:
                    ref_act = respuestas_efectivas(arch)
                    st.markdown("---")
                    st.markdown("**Detalle de respuestas:**")
                    html="<div style='line-height:2.4;'>"
                    for i,(r,p) in enumerate(zip(ref_act,pauta),1):
                        es_dud  = i in dudosas
                        fue_c   = str(i) in st.session_state.correcciones.get(arch,{}) and es_dud
                        if r is None:
                            cls="badge-null"; txt=f"P{i}:—"
                        elif p and r.upper()==p.upper():
                            cls="badge-ok";  txt=f"P{i}:{r}"
                        else:
                            cls="badge-err"; txt=f"P{i}:{r or '?'}"
                        if es_dud and not fue_c: cls="badge-dud"
                        html+=f'<span class="corr-badge {cls}">{txt}{"✎" if fue_c else ""}</span>'
                    html+="</div>"
                    st.markdown(html, unsafe_allow_html=True)


# ══ EXPORTAR ═════════════════════════════════════════════════════════
with tab_exportar:
    pauta = st.session_state["pauta"]

    if not st.session_state.resultados:
        st.info("Procesa al menos una hoja antes de exportar.")
    elif not any(pauta):
        st.error("Ve a **Pauta de respuestas** e ingresa las respuestas correctas primero.")
    else:
        pend_exp = sum(
            1 for a,d in st.session_state.resultados.items()
            if [x for x in d.get("dudosas",[]) if str(x) not in st.session_state.correcciones.get(a,{})]
        )
        total_p=sum(1 for p in pauta if p)
        sospechosas_exp = sum(1 for d in st.session_state.resultados.values() if d.get("sospechoso"))
        sin_id_exp = sum(
            1 for a in st.session_state.resultados
            if not (datos_efectivos(a).get("apellido_paterno") or datos_efectivos(a).get("nombres")
                     or datos_efectivos(a).get("cedula"))
        )
        m1,m2,m3,m4,m5=st.columns(5)
        m1.metric("Alumnos",len(st.session_state.resultados))
        m2.metric("Preguntas evaluadas",total_p)
        m3.metric("Dudas pendientes",pend_exp,
                  delta="revisar antes" if pend_exp else "todo resuelto",
                  delta_color="inverse" if pend_exp else "normal")
        m4.metric("🚨 Sospechosas",sospechosas_exp,
                  delta="revisar antes" if sospechosas_exp else "ninguna",
                  delta_color="inverse" if sospechosas_exp else "normal")
        m5.metric("🆔 Sin identificar",sin_id_exp,
                  delta="completar antes" if sin_id_exp else "todo identificado",
                  delta_color="inverse" if sin_id_exp else "normal")

        if sospechosas_exp:
            st.error(f"🚨 {sospechosas_exp} hoja(s) con patrón de respuestas estadísticamente inverosímil — "
                     "revísalas en **Revisar y corregir** antes de confiar en el Excel. Quedan marcadas en la "
                     "columna 'Alerta calidad' del Resumen igualmente.")
        if sin_id_exp:
            st.warning(f"🆔 {sin_id_exp} hoja(s) sin apellido, nombre ni cédula — el Excel las mostrará en blanco "
                       "en esas columnas. Complétalas en **Revisar y corregir** antes de exportar si necesitas "
                       "identificar a cada alumno.")
        if pend_exp:
            st.warning(f"⚠️ {pend_exp} alumno(s) con dudas sin corregir — quedarán en amarillo.")
        elif not sospechosas_exp and not sin_id_exp:
            st.success("✅ Todo revisado. El Excel estará completo.")

        nombre_xl=f"resultados_{curso.replace(' ','_') if curso else 'curso'}.xlsx"
        excel=generar_excel(pauta, curso)
        st.download_button(
            label="⬇️ Descargar Excel", data=excel,
            file_name=nombre_xl,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary", use_container_width=True,
        )
        st.markdown("---")
        st.markdown("""
El Excel incluye 4 hojas:
- **Resumen** — puntaje y % de logro por alumno, con columna **Curso** en cada fila:
  puedes copiar y pegar las filas de varios Excel descargados en un mismo archivo
  maestro y luego filtrar/ordenar por Curso o por alumno (Ap. Paterno / Nombres).
- **Detalle respuestas** — cada pregunta en verde ✅ / rojo ❌ / amarillo ⚠️
- **Estadísticas del curso** — promedio, máximo, mínimo y distribución
- **Pauta utilizada** — registro de las respuestas correctas usadas

**Recomendación:** procesa las fotos en lotes de ~15–20 antes de exportar,
en vez de subir 40–80 de una vez. Puedes exportar el Excel parcial de cada
lote y luego pegarlas todas en el archivo maestro — así no pierdes avance
si el celular se bloquea o la conexión se corta a mitad de un lote grande.
""")
