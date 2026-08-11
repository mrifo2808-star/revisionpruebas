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
import qrcode
import pandas as pd
import streamlit as st
import anthropic
from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Revisor de Pruebas",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="expanded",
)

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
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

OPCIONES = ["A", "B", "C", "D", "E", "—"]


def prompt_dinamico(n: int) -> str:
    return f"""Eres un asistente experto en corregir hojas de respuestas de alternativas de estudiantes chilenos.

PASO 1 — Entiende el layout antes de leer ninguna marca:
Muchas hojas de respuestas chilenas NO son una sola lista de arriba hacia abajo: están organizadas en 2, 3 o
4 columnas de preguntas puestas una al lado de la otra (por ejemplo, preguntas 1-20 en la primera columna,
21-40 en la segunda, 41-60 en la tercera, etc). Antes de leer ninguna marca, identifica cuántas columnas de
preguntas tiene esta hoja específica y en qué número empieza y termina cada una.

PASO 2 — Usa el número IMPRESO, nunca el orden de lectura:
Cada fila tiene un número de pregunta impreso junto a las burbujas A-E (ej. "21 (A)(B)(C)(D)(E)"). Ese número
impreso es la única fuente de verdad sobre a qué pregunta corresponde esa fila — NO asumas que la fila que ves
"más abajo" o "en la posición 21 según fuiste mirando" es la pregunta 21. Verifica el número impreso de cada
fila antes de registrar su respuesta.

PASO 3 — Registra cada respuesta en el índice correcto del arreglo:
El arreglo "respuestas" tiene {n} posiciones, donde la posición 1 = pregunta impresa "1", la posición 2 =
pregunta impresa "2", y así sucesivamente — sin importar en qué orden espacial las hayas ido mirando en la
imagen (por columnas, no de arriba a abajo en un solo bloque).

PASO 4 — Verifica antes de responder:
Revisa que el arreglo "respuestas" tenga exactamente {n} elementos y que corresponda una posición por cada
número de pregunta impreso en la hoja, del 1 al {n}, sin saltos ni desplazamientos. Luego, vuelve a mirar por
segunda vez SOLO las preguntas donde no quedaste 100% seguro de cuál opción marcó el estudiante, y confírmalas
con calma.

Criterio simple para "dudosas" — marca una pregunta como dudosa ÚNICAMENTE si, tras esa segunda mirada, sigue
existiendo un riesgo real de haber leído mal la intención del estudiante (ejemplos: dos opciones con marca
igual de oscura, un borrón que deja la burbuja ambigua entre dos letras, o una marca tan tenue que pudo no
ser intencional). Si el trazo es imperfecto o desprolijo pero al mirarlo con calma se entiende con claridad
cuál opción eligió el estudiante, NO es una pregunta dudosa — entrégala como respuesta normal, sin marcarla.

El objetivo es que "dudosas" quede lo más corta posible y contenga solo los casos con riesgo real de error;
todo lo demás se da por bueno sin necesitar revisión humana.

Responde ÚNICAMENTE con un JSON válido, sin texto adicional ni markdown, con esta forma exacta:

{{
  "apellido_paterno": "...",
  "apellido_materno": "...",
  "nombres": "...",
  "cedula": "...",
  "nro_folleto": "...",
  "respuestas": ["A","B",...],
  "dudosas": [3, 15]
}}

Reglas de formato:
- "respuestas": exactamente {n} elementos, en el mismo orden P1..P{n}. Usa null solo si la pregunta está
  realmente omitida (ninguna marca visible), no como comodín para lo que no estés seguro.
- "dudosas": números de pregunta (1 a {n}) con riesgo real de error, según el criterio simple de arriba.
- Campos de texto ilegibles o no visibles en la hoja: cadena vacía "".
- Solo el JSON, sin explicación ni comentarios adicionales."""


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

def comprimir_imagen(datos_bytes: bytes, mime: str, lado_max: int = 1600, calidad: int = 85):
    """Reduce tamaño/resolución para que la subida desde datos móviles no se cuelgue."""
    try:
        img = Image.open(io.BytesIO(datos_bytes))
        img = img.convert("RGB")
        if max(img.size) > lado_max:
            escala = lado_max / max(img.size)
            img = img.resize((int(img.width*escala), int(img.height*escala)), Image.Resampling.LANCZOS)
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=calidad, optimize=True)
        return buf.getvalue(), "image/jpeg"
    except Exception:
        return datos_bytes, mime

def procesar_imagen(cliente, nombre: str, datos_bytes: bytes, mime: str, n: int) -> dict:
    data = base64.standard_b64encode(datos_bytes).decode()
    msg = cliente.messages.create(
        model="claude-sonnet-5", max_tokens=4096,
        messages=[{"role":"user","content":[
            {"type":"image","source":{"type":"base64","media_type":mime,"data":data}},
            {"type":"text","text":prompt_dinamico(n)},
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
    res["archivo"] = nombre
    res["n_preguntas"] = n
    resp = res.get("respuestas", [])
    res["respuestas"] = (resp + [None]*n)[:n]
    res["dudosas"] = sorted(set(res.get("dudosas", [])))
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
           "Preguntas incorrectas","Dudosas corregidas"]
    COL_PUNTAJE = 11
    COLS_TEXTO = {1,3,4,5,6,7}
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
        dud = st.session_state.resultados[arch].get("dudosas",[])
        corr= st.session_state.correcciones.get(arch,{})
        cn  = len([k for k in corr if int(k) in dud])
        vals=[curso or "—", rn,
              datos.get("apellido_paterno",""),datos.get("apellido_materno",""),
              datos.get("nombres",""),datos.get("cedula",""),datos.get("nro_folleto",""),
              co,inc,om,pct,
              ", ".join(str(e) for e in err) if err else "—",
              f"{cn} de {len(dud)}" if dud else "—"]
        for c,v in enumerate(vals,1):
            cell=ws1.cell(fe+rn,c,v); cell.border=bd
            cell.alignment=Alignment(horizontal="left" if c in COLS_TEXTO else "center",wrap_text=True)
            if c==COL_PUNTAJE: cell.fill=VE if pct>=70 else (AM if pct>=50 else RO)
    for i,w in enumerate([20,4,16,16,22,14,9,10,11,10,10,38,18],1):
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
    if st.button("🗑️ Limpiar todo", use_container_width=True):
        nn = st.session_state["n_preguntas"]
        st.session_state.update({
            "resultados":{}, "correcciones":{}, "info_edits":{},
            "fotos_pendientes":{}, "pauta":[], "pauta_df":df_pauta_vacio(nn),
        })
        st.rerun()


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
    st.caption("Haz clic en una celda de **Respuesta** y elige A–E. O usa importación rápida a la derecha.")
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
            use_container_width=True, key="editor_pauta",
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

    if app_url:
        cq, ci = st.columns([1,2], gap="large")
        with cq:
            st.markdown("### 📱 Desde el celular")
            st.image(generar_qr(app_url), width=180, caption="Escanea para abrir en tu celular")
        with ci:
            st.markdown("### Instrucciones")
            st.markdown("""
1. Escanea el QR o entra a la URL en el navegador del celular
2. Toca **Cargar fotos** arriba
3. Toca **Upload** → elige **Cámara** o **Galería**
4. Buena iluminación, encuadra toda la hoja
5. Sube varias antes de procesar
""")
    else:
        st.info("Desde el celular: entra a la URL de esta app en el navegador. "
                "Al tocar Upload el celular ofrece abrir la cámara directamente. "
                "Agrega `APP_URL` a Secrets de Streamlit Cloud para ver un QR aquí.")

    st.markdown("---")
    st.markdown(f"### Sube las fotos  ·  *{n} preguntas por prueba*")
    st.caption("En el celular puedes tocar el recuadro varias veces para tomar una foto a la vez: "
               "cada una queda guardada aunque la cámara se abra de nuevo.")
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
                comprimido, mime_final = comprimir_imagen(contenido, TIPOS_MIME.get(f.type, "image/jpeg"))
                idx = len(st.session_state.fotos_pendientes) + len(st.session_state.resultados) + 1
                st.session_state.fotos_pendientes[f"foto_{idx:03d}"] = {
                    "nombre": f.name, "bytes": comprimido,
                    "mime": mime_final, "hash": h,
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
            cliente = anthropic.Anthropic(api_key=key, timeout=60.0, max_retries=1)
            prog = st.progress(0, text="Iniciando...")
            errs = []
            items = list(pendientes.items())
            for i, (id_unico, foto) in enumerate(items):
                prog.progress(i/len(items), text=f"Procesando {foto['nombre']} ({i+1}/{len(items)})...")
                try:
                    res = procesar_imagen(cliente, foto["nombre"], foto["bytes"], foto["mime"], n)
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
        m1,m2,m3 = st.columns(3)
        m1.metric("Alumnos procesados", total_alumnos)
        m2.metric("Con dudas pendientes", pendientes)
        m3.metric("Preguntas por prueba", n)
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

            icono = "⚠️" if pend_este else "✅"
            with st.expander(
                f"{icono} {datos.get('apellido_paterno','')} {datos.get('nombres','')} "
                f"— {datos.get('cedula','')} | {puntaje_str}",
                expanded=bool(pend_este)
            ):
                # ── Datos editables del alumno ──────────────────────
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
        m1,m2,m3=st.columns(3)
        m1.metric("Alumnos",len(st.session_state.resultados))
        m2.metric("Preguntas evaluadas",total_p)
        m3.metric("Dudas pendientes",pend_exp,
                  delta="revisar antes" if pend_exp else "todo resuelto",
                  delta_color="inverse" if pend_exp else "normal")

        if pend_exp:
            st.warning(f"⚠️ {pend_exp} alumno(s) con dudas sin corregir — quedarán en amarillo.")
        else:
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
