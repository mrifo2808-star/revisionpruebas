"""
Revisor de Hojas de Respuestas — App Streamlit
Versión con pauta numerada y soporte para nube (Streamlit Community Cloud)
"""

import io
import json
import re
import base64
import pandas as pd
import streamlit as st
import anthropic
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─── Configuración ───────────────────────────────────────────────────
st.set_page_config(
    page_title="Revisor de Pruebas",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
.corr-badge {
    display:inline-block; padding:2px 8px; border-radius:6px;
    font-size:12px; font-weight:600; margin:2px;
}
.badge-ok   { background:#dcfce7; color:#166534; }
.badge-err  { background:#fee2e2; color:#991b1b; }
.badge-dud  { background:#fef9c3; color:#92400e; }
.badge-null { background:#f3f4f6; color:#6b7280; }
div[data-testid="stDataEditor"] { font-size: 13px; }
</style>
""", unsafe_allow_html=True)

# ─── Estado de sesión ────────────────────────────────────────────────
defaults = {
    "resultados": {},
    "correcciones": {},
    "pauta": [],
    "procesados": set(),
    "pauta_df": pd.DataFrame({
        "N°": [f"P{i}" for i in range(1, 81)],
        "Respuesta": [""] * 80,
    }),
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

OPCIONES = ["A", "B", "C", "D", "E", "—"]

PROMPT = """Analiza esta hoja de respuestas de un estudiante chileno.

Responde ÚNICAMENTE con un JSON válido, sin texto adicional ni markdown:

{
  "apellido_paterno": "...",
  "apellido_materno": "...",
  "nombres": "...",
  "cedula": "...",
  "nro_folleto": "...",
  "respuestas": ["A","B",...],
  "dudosas": [3, 15, 42]
}

Reglas:
- "respuestas": exactamente 80 elementos en orden. null si burbuja no marcada o ilegible.
- "dudosas": números de pregunta (1-80) con burbuja poco marcada, borrada o ambigua.
  Si hay dos marcadas, elige la más oscura y de todas formas ponla en "dudosas".
- Campos de texto ilegibles: cadena vacía "".
- Solo JSON, sin explicación."""


# ─── Funciones ───────────────────────────────────────────────────────

def api_key_activa() -> str:
    """Devuelve la API key desde secrets (nube) o desde el input del usuario."""
    try:
        return st.secrets["ANTHROPIC_API_KEY"]
    except Exception:
        return st.session_state.get("api_key_input", "")


def imagen_a_base64(uploaded_file):
    tipos = {"image/jpeg": "image/jpeg", "image/jpg": "image/jpeg",
             "image/png": "image/png", "image/webp": "image/webp"}
    media_type = tipos.get(uploaded_file.type, "image/jpeg")
    data = base64.standard_b64encode(uploaded_file.read()).decode()
    uploaded_file.seek(0)
    return data, media_type


def procesar_imagen(cliente, uploaded_file) -> dict:
    data, media_type = imagen_a_base64(uploaded_file)
    msg = cliente.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=2048,
        messages=[{"role": "user", "content": [
            {"type": "image", "source": {"type": "base64",
             "media_type": media_type, "data": data}},
            {"type": "text", "text": PROMPT},
        ]}],
    )
    texto = msg.content[0].text.strip()
    match = re.search(r'\{[\s\S]*\}', texto)
    if match:
        texto = match.group(0)
    resultado = json.loads(texto)
    resultado["archivo"] = uploaded_file.name
    resp = resultado.get("respuestas", [])
    resultado["respuestas"] = (resp + [None] * 80)[:80]
    resultado["dudosas"] = sorted(set(resultado.get("dudosas", [])))
    return resultado


def respuestas_efectivas(nombre_archivo: str) -> list:
    base = st.session_state.resultados[nombre_archivo]["respuestas"].copy()
    for idx_str, letra in st.session_state.correcciones.get(nombre_archivo, {}).items():
        base[int(idx_str) - 1] = None if letra == "—" else letra
    return base


def calcular(respuestas, pauta):
    correctas = incorrectas = omitidas = 0
    errores = []
    for i, (r, p) in enumerate(zip(respuestas, pauta), 1):
        if not p:
            continue
        if r is None:
            omitidas += 1
        elif r.upper() == p.upper():
            correctas += 1
        else:
            incorrectas += 1
            errores.append(i)
    return correctas, incorrectas, omitidas, errores


def pauta_desde_df(df: pd.DataFrame) -> list:
    return [
        r.strip().upper() if r and r.strip().upper() in "ABCDE" else None
        for r in df["Respuesta"].fillna("").tolist()
    ]


def generar_excel(pauta: list, curso: str) -> bytes:
    wb = Workbook()
    borde = Border(left=Side(style="thin"), right=Side(style="thin"),
                   top=Side(style="thin"), bottom=Side(style="thin"))
    AZ  = PatternFill("solid", fgColor="2563EB")
    AZD = PatternFill("solid", fgColor="1E3A5F")
    VE  = PatternFill("solid", fgColor="DCFCE7")
    RO  = PatternFill("solid", fgColor="FECACA")
    AM  = PatternFill("solid", fgColor="FEF08A")
    AZC = PatternFill("solid", fgColor="DBEAFE")
    GR  = PatternFill("solid", fgColor="F3F4F6")
    total_p = sum(1 for p in pauta if p)

    # Hoja 1 — Resumen
    ws1 = wb.active
    ws1.title = "Resumen"
    if curso:
        ws1["A1"] = f"Curso: {curso}"
        ws1["A1"].font = Font(bold=True, size=13)
        ws1.merge_cells("A1:M1")

    enc = ["N°","Archivo","Ap. Paterno","Ap. Materno","Nombres","Cédula",
           "Folleto","Correctas","Incorrectas","Omitidas","Puntaje %",
           "Preguntas incorrectas","Dudosas corregidas"]
    fe = 3
    for c, h in enumerate(enc, 1):
        cell = ws1.cell(fe, c, h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = AZ
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = borde

    for n, (arch, datos) in enumerate(st.session_state.resultados.items(), 1):
        ref = respuestas_efectivas(arch)
        corr = st.session_state.correcciones.get(arch, {})
        co, inc, om, err = calcular(ref, pauta)
        pct = round(co / total_p * 100, 1) if total_p else 0
        dud = datos.get("dudosas", [])
        corr_n = len([k for k in corr if int(k) in dud])
        vals = [n, arch,
                datos.get("apellido_paterno",""), datos.get("apellido_materno",""),
                datos.get("nombres",""), datos.get("cedula",""),
                datos.get("nro_folleto",""),
                co, inc, om, pct,
                ", ".join(str(e) for e in err) if err else "—",
                f"{corr_n} de {len(dud)}" if dud else "—"]
        for c, v in enumerate(vals, 1):
            cell = ws1.cell(fe + n, c, v)
            cell.border = borde
            cell.alignment = Alignment(
                horizontal="center" if c not in (2,3,4,5,6) else "left",
                wrap_text=True)
            if c == 11:
                cell.fill = VE if pct >= 70 else (AM if pct >= 50 else RO)

    for i, w in enumerate([4,22,15,15,22,14,9,10,11,10,10,38,18], 1):
        ws1.column_dimensions[get_column_letter(i)].width = w
    ws1.row_dimensions[fe].height = 32

    # Hoja 2 — Detalle respuestas
    ws2 = wb.create_sheet("Detalle respuestas")
    cab = ["Alumno"] + [f"P{i}" for i in range(1, 81)]
    for c, h in enumerate(cab, 1):
        cell = ws2.cell(1, c, h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = AZD
        cell.alignment = Alignment(horizontal="center")
        cell.border = borde

    ws2.cell(2, 1, "PAUTA").font = Font(bold=True)
    for i, p in enumerate(pauta, 2):
        cell = ws2.cell(2, i, p or "?")
        cell.fill = AZC
        cell.alignment = Alignment(horizontal="center")
        cell.border = borde

    for fila, (arch, datos) in enumerate(st.session_state.resultados.items(), 3):
        ref = respuestas_efectivas(arch)
        corr = st.session_state.correcciones.get(arch, {})
        dud_set = set(datos.get("dudosas", []))
        nombre = f"{datos.get('apellido_paterno','')} {datos.get('nombres','')}"
        ws2.cell(fila, 1, nombre).border = borde
        for i, r in enumerate(ref, 1):
            cell = ws2.cell(fila, i+1, r or "—")
            cell.alignment = Alignment(horizontal="center")
            cell.border = borde
            corregida = str(i) in corr and i in dud_set
            if r is None:
                cell.fill = GR
            elif pauta[i-1] and r.upper() == pauta[i-1].upper():
                cell.fill = VE
                if corregida:
                    cell.font = Font(bold=True, color="166534")
            else:
                cell.fill = RO
                if corregida:
                    cell.font = Font(bold=True, color="991B1B")
            if i in dud_set and not corregida:
                cell.fill = AM
                cell.font = Font(bold=True, color="92400E")

    ws2.column_dimensions["A"].width = 26
    for col in range(2, 82):
        ws2.column_dimensions[get_column_letter(col)].width = 4.2
    ws2.row_dimensions[1].height = 24

    # Hoja 3 — Estadísticas
    ws3 = wb.create_sheet("Estadísticas")
    ws3["A1"] = f"Estadísticas — {curso or 'Curso'}"
    ws3["A1"].font = Font(bold=True, size=14)
    todos_pct = []
    for arch in st.session_state.resultados:
        ref = respuestas_efectivas(arch)
        co, _, _, _ = calcular(ref, pauta)
        todos_pct.append(round(co / total_p * 100, 1) if total_p else 0)
    stats = [
        ("Total alumnos", len(todos_pct)),
        ("Promedio del curso (%)", round(sum(todos_pct)/len(todos_pct), 1) if todos_pct else 0),
        ("Puntaje máximo (%)", max(todos_pct) if todos_pct else 0),
        ("Puntaje mínimo (%)", min(todos_pct) if todos_pct else 0),
        ("Sobre 60% de logro", sum(1 for p in todos_pct if p >= 60)),
        ("Bajo 60% de logro", sum(1 for p in todos_pct if p < 60)),
        ("Hojas con dudas pendientes", sum(
            1 for a, d in st.session_state.resultados.items()
            if [x for x in d.get("dudosas",[]) if str(x) not in st.session_state.correcciones.get(a,{})]
        )),
    ]
    for f, (e, v) in enumerate(stats, 3):
        ws3.cell(f, 1, e).font = Font(bold=True)
        ws3.cell(f, 2, v)
    ws3.column_dimensions["A"].width = 34
    ws3.column_dimensions["B"].width = 14

    # Hoja 4 — Pauta usada
    ws4 = wb.create_sheet("Pauta utilizada")
    ws4["A1"] = "Pauta de respuestas usada para este curso"
    ws4["A1"].font = Font(bold=True, size=12)
    ws4.merge_cells("A1:D1")
    for i in range(0, 80, 4):
        for j, idx in enumerate(range(i, min(i+4, 80))):
            col = j * 2 + 1
            fila = i // 4 + 3
            ws4.cell(fila, col, f"P{idx+1}").font = Font(bold=True, color="1E3A5F")
            cell = ws4.cell(fila, col+1, pauta[idx] or "—")
            cell.alignment = Alignment(horizontal="center")
            if pauta[idx]:
                cell.fill = AZC
    for col in range(1, 9):
        ws4.column_dimensions[get_column_letter(col)].width = 7

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ═══════════════════════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## ⚙️ Configuración")

    # API Key: desde secrets (nube) o manual
    tiene_secret = False
    try:
        _ = st.secrets["ANTHROPIC_API_KEY"]
        tiene_secret = True
    except Exception:
        pass

    if tiene_secret:
        st.success("🔑 API Key cargada desde configuración de nube")
    else:
        st.session_state["api_key_input"] = st.text_input(
            "API Key de Anthropic", type="password",
            placeholder="sk-ant-...",
            help="Obtén tu clave en console.anthropic.com")

    curso = st.text_input("Nombre del curso", placeholder="Ej: 3°A — Historia 2026")

    st.markdown("---")
    if st.button("🗑️ Limpiar todo", use_container_width=True):
        for k in ["resultados", "correcciones", "procesados"]:
            st.session_state[k] = {} if k != "procesados" else set()
        st.session_state["pauta_df"] = pd.DataFrame({
            "N°": [f"P{i}" for i in range(1, 81)],
            "Respuesta": [""] * 80,
        })
        st.session_state["pauta"] = []
        st.rerun()


# ═══════════════════════════════════════════════════════════════════════
# CUERPO PRINCIPAL
# ═══════════════════════════════════════════════════════════════════════
st.markdown("# 📝 Revisor de Hojas de Respuestas")

tab_pauta, tab_cargar, tab_revisar, tab_exportar = st.tabs([
    "📋 Pauta de respuestas",
    "📤 Cargar y procesar",
    "✏️ Revisar y corregir",
    "📊 Exportar Excel",
])

# ══ TAB 0: PAUTA ════════════════════════════════════════════════════
with tab_pauta:
    st.markdown("### Ingresa las respuestas correctas por número de pregunta")
    st.caption("Haz clic en cualquier celda de la columna **Respuesta** y elige A, B, C, D o E.")

    col_tabla, col_import = st.columns([2, 1], gap="large")

    with col_import:
        st.markdown("**Importar rápido**")
        st.caption("Pega todas las respuestas en orden (una por línea o separadas por coma/espacio):")
        texto_bulk = st.text_area("Respuestas en bloque", height=180,
            placeholder="A\nB\nC\n...\no bien: A,B,C,D,E,...",
            label_visibility="collapsed")

        if st.button("⬇️ Cargar al grid", use_container_width=True):
            letras = re.findall(r'[AaBbCcDdEe]', texto_bulk)
            letras = [l.upper() for l in letras]
            if letras:
                nuevas = (letras + [""] * 80)[:80]
                st.session_state["pauta_df"] = pd.DataFrame({
                    "N°": [f"P{i}" for i in range(1, 81)],
                    "Respuesta": nuevas,
                })
                st.success(f"✓ {len([l for l in letras if l])} respuestas cargadas")
                st.rerun()
            else:
                st.error("No se encontraron letras A–E válidas.")

        st.markdown("---")
        pauta_actual = pauta_desde_df(st.session_state["pauta_df"])
        completadas = sum(1 for p in pauta_actual if p)
        st.metric("Preguntas con respuesta", f"{completadas} / 80")
        if completadas == 80:
            st.success("✅ Pauta completa")
        elif completadas > 0:
            st.warning(f"⚠️ Faltan {80 - completadas} respuestas")

    with col_tabla:
        df_editado = st.data_editor(
            st.session_state["pauta_df"],
            column_config={
                "N°": st.column_config.TextColumn(
                    "N°", disabled=True, width="small"),
                "Respuesta": st.column_config.SelectboxColumn(
                    "Respuesta",
                    options=["A", "B", "C", "D", "E"],
                    width="small",
                    required=False,
                ),
            },
            hide_index=True,
            height=520,
            use_container_width=True,
            key="editor_pauta",
        )
        # Guardar cambios del editor en session state
        st.session_state["pauta_df"] = df_editado
        st.session_state["pauta"] = pauta_desde_df(df_editado)

    # Vista previa compacta
    pauta = st.session_state["pauta"]
    if any(pauta):
        with st.expander("👁️ Vista previa de la pauta cargada"):
            html = "<div style='font-family:monospace; font-size:13px; line-height:2.2;'>"
            for i, p in enumerate(pauta, 1):
                color = "#dcfce7" if p else "#f3f4f6"
                txt_color = "#166534" if p else "#9ca3af"
                html += (f'<span style="display:inline-block;width:52px;margin:2px;'
                         f'background:{color};color:{txt_color};border-radius:4px;'
                         f'padding:2px 4px;font-size:12px;font-weight:600;">'
                         f'P{i}:{p or "?"}</span>')
            html += "</div>"
            st.markdown(html, unsafe_allow_html=True)


# ══ TAB 1: CARGAR IMÁGENES ══════════════════════════════════════════
with tab_cargar:
    st.markdown("### Sube las fotos de las hojas de respuestas")
    st.caption("Se aceptan JPG, PNG y WEBP. Puedes subir varias a la vez.")

    archivos = st.file_uploader(
        "Fotos",
        type=["jpg", "jpeg", "png", "webp"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

    if archivos:
        nuevos = [f for f in archivos if f.name not in st.session_state.procesados]
        ya = [f for f in archivos if f.name in st.session_state.procesados]

        if ya:
            st.info(f"✓ {len(ya)} hoja(s) ya procesadas.")
        if nuevos:
            st.warning(f"⏳ {len(nuevos)} hoja(s) nueva(s) lista(s) para procesar.")
            key = api_key_activa()
            if not key:
                st.error("Ingresa tu API Key en el panel lateral para continuar.")
            else:
                if st.button(f"🚀 Procesar {len(nuevos)} hoja(s)", type="primary",
                             use_container_width=True):
                    cliente = anthropic.Anthropic(api_key=key)
                    progreso = st.progress(0, text="Iniciando...")
                    errores_proc = []
                    for i, archivo in enumerate(nuevos):
                        progreso.progress(i / len(nuevos),
                            text=f"Procesando {archivo.name} ({i+1}/{len(nuevos)})...")
                        try:
                            res = procesar_imagen(cliente, archivo)
                            st.session_state.resultados[archivo.name] = res
                            st.session_state.procesados.add(archivo.name)
                        except Exception as e:
                            errores_proc.append(f"{archivo.name}: {e}")
                    progreso.progress(1.0, text="¡Completado!")
                    if errores_proc:
                        st.error("Errores:\n" + "\n".join(errores_proc))
                    else:
                        st.success(f"✅ {len(nuevos)} hoja(s) procesadas. Ve a **Revisar y corregir**.")
                    st.rerun()
        else:
            st.success(f"✅ Todas las hojas cargadas ({len(archivos)} en total).")
    else:
        st.markdown("""
        <div style="border:2px dashed #d1d5db;border-radius:12px;padding:2rem;
                    text-align:center;color:#6b7280;">
            <div style="font-size:2.5rem;">📄</div>
            <div style="font-size:15px;margin-top:8px;">Arrastra aquí las fotos</div>
            <div style="font-size:13px;margin-top:4px;">JPG · PNG · WEBP · hasta 200 MB por archivo</div>
        </div>""", unsafe_allow_html=True)


# ══ TAB 2: REVISAR Y CORREGIR ═══════════════════════════════════════
with tab_revisar:
    pauta = st.session_state["pauta"]

    if not st.session_state.resultados:
        st.info("Aún no hay hojas procesadas. Ve a **Cargar y procesar**.")
    else:
        total = len(st.session_state.resultados)
        pendientes = sum(
            1 for a, d in st.session_state.resultados.items()
            if [x for x in d.get("dudosas",[])
                if str(x) not in st.session_state.correcciones.get(a,{})]
        )
        c1, c2, c3 = st.columns(3)
        c1.metric("Alumnos", total)
        c2.metric("Con dudas pendientes", pendientes,
                  delta=f"-{total-pendientes} resueltos" if total > pendientes else None,
                  delta_color="normal")
        c3.metric("Preguntas en pauta", sum(1 for p in pauta if p))
        st.markdown("---")

        for arch, datos in st.session_state.resultados.items():
            apellido = datos.get("apellido_paterno","")
            nombres  = datos.get("nombres","")
            cedula   = datos.get("cedula","")
            dudosas  = datos.get("dudosas",[])
            corr_arch = st.session_state.correcciones.get(arch, {})
            pendientes_este = [d for d in dudosas if str(d) not in corr_arch]

            ref = respuestas_efectivas(arch)
            if pauta:
                co, inc, om, _ = calcular(ref, pauta)
                total_p = sum(1 for p in pauta if p)
                pct = round(co / total_p * 100, 1) if total_p else 0
                puntaje_str = f"**{pct}%** ({co}/{total_p})"
            else:
                puntaje_str = "_(sin pauta)_"

            icono = "⚠️" if pendientes_este else "✅"
            with st.expander(
                f"{icono} {apellido} {nombres} — {cedula} | {puntaje_str}",
                expanded=bool(pendientes_este)
            ):
                ca, cb, cc, cd = st.columns(4)
                ca.markdown(f"**Ap. Paterno:** {datos.get('apellido_paterno','—')}")
                cb.markdown(f"**Ap. Materno:** {datos.get('apellido_materno','—')}")
                cc.markdown(f"**Nombres:** {datos.get('nombres','—')}")
                cd.markdown(f"**Cédula:** {datos.get('cedula','—')}")

                if dudosas:
                    st.markdown(f"**Preguntas dudosas:** {', '.join(f'P{d}' for d in dudosas)}")
                    st.markdown("Selecciona la respuesta correcta para cada una:")
                    n_cols = min(len(dudosas), 8)
                    cols = st.columns(n_cols)
                    for j, num_p in enumerate(dudosas):
                        resp_orig = datos["respuestas"][num_p - 1]
                        resp_actual = corr_arch.get(str(num_p), resp_orig or "—")
                        with cols[j % n_cols]:
                            nueva = st.selectbox(
                                f"P{num_p}",
                                options=OPCIONES,
                                index=OPCIONES.index(resp_actual) if resp_actual in OPCIONES else 5,
                                key=f"sel_{arch}_{num_p}",
                            )
                            if nueva != (resp_orig or "—"):
                                if arch not in st.session_state.correcciones:
                                    st.session_state.correcciones[arch] = {}
                                st.session_state.correcciones[arch][str(num_p)] = nueva
                            elif str(num_p) in st.session_state.correcciones.get(arch, {}):
                                del st.session_state.correcciones[arch][str(num_p)]
                else:
                    st.success("Sin preguntas dudosas.")

                if pauta:
                    st.markdown("**Detalle de respuestas:**")
                    ref_actual = respuestas_efectivas(arch)
                    html = "<div style='line-height:2;'>"
                    for i, (r, p) in enumerate(zip(ref_actual, pauta), 1):
                        es_dud = i in dudosas
                        fue_corr = str(i) in corr_arch and es_dud
                        if r is None:
                            cls = "badge-null"; txt = f"P{i}:—"
                        elif p and r.upper() == p.upper():
                            cls = "badge-ok"; txt = f"P{i}:{r}"
                        else:
                            cls = "badge-err"; txt = f"P{i}:{r or '?'}"
                        if es_dud and not fue_corr:
                            cls = "badge-dud"
                        html += f'<span class="corr-badge {cls}">{txt}{"✎" if fue_corr else ""}</span>'
                    html += "</div>"
                    st.markdown(html, unsafe_allow_html=True)


# ══ TAB 3: EXPORTAR ═════════════════════════════════════════════════
with tab_exportar:
    pauta = st.session_state["pauta"]

    if not st.session_state.resultados:
        st.info("Procesa al menos una hoja antes de exportar.")
    elif not any(pauta):
        st.error("Ve a la pestaña **Pauta de respuestas** e ingresa las respuestas correctas.")
    else:
        pendientes_exp = sum(
            1 for a, d in st.session_state.resultados.items()
            if [x for x in d.get("dudosas",[])
                if str(x) not in st.session_state.correcciones.get(a,{})]
        )
        if pendientes_exp:
            st.warning(f"⚠️ Hay **{pendientes_exp} alumno(s)** con dudas sin corregir. "
                       "Aparecerán en amarillo en el Excel.")
        else:
            st.success("✅ Todas las dudas resueltas. El Excel estará completo.")

        nombre_xl = f"resultados_{curso.replace(' ','_') if curso else 'curso'}.xlsx"
        excel = generar_excel(pauta, curso)

        st.download_button(
            label="⬇️ Descargar Excel",
            data=excel,
            file_name=nombre_xl,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
        st.markdown("---")
        st.markdown("""
El Excel incluye:
- **Resumen** — puntaje y % de logro por alumno, con colores según rendimiento
- **Detalle respuestas** — cada pregunta en verde ✅ / rojo ❌ / amarillo ⚠️
- **Estadísticas del curso** — promedio, máximo, mínimo y distribución
- **Pauta utilizada** — registro de las respuestas correctas usadas
""")
