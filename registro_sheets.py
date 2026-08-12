"""
Cliente de Google Sheets para la plataforma "Herramientas Docentes":

- Registro de USUARIOS (login) en la pestaña "usuarios" -- grupo cerrado,
  cuentas precargadas por el administrador (ver gestionar_usuarios.py),
  sin auto-registro público.
- Registro de ACTIVIDAD (respaldo de solicitudes/resultados entregados) en
  la pestaña "registro" -- SOLO METADATOS (quién, cuándo, curso, cantidad
  de hojas y contadores de confiables/dudosas). Nunca nombres, RUT ni
  respuestas de estudiantes: eso sigue viviendo únicamente en el Excel que
  el docente descarga, no en este respaldo.

Streamlit Community Cloud no garantiza persistir archivos locales entre
reinicios del contenedor -- por eso el respaldo vive en una hoja de cálculo
externa, no en un archivo local.

Credenciales esperadas en st.secrets (mismo patrón que ya usa
api_key_activa()/tiene_secret() en app_revisor.py para la API key de
Anthropic -- nunca committeadas al repo, .streamlit/secrets.toml está en
.gitignore):

    GOOGLE_SHEETS_ID = "el-id-de-la-planilla"
    [gcp_service_account]
    type = "service_account"
    project_id = "..."
    private_key_id = "..."
    private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
    client_email = "...@....iam.gserviceaccount.com"
    client_id = "..."
    token_uri = "https://oauth2.googleapis.com/token"

Además (para "recordarme" en el login, ver app_revisor.py):

    AUTH_COOKIE_SECRET = "una-cadena-larga-y-aleatoria-cualquiera"

Sin AUTH_COOKIE_SECRET, "recordarme" queda deshabilitado sin romper el login
normal (fail-closed) -- generarla una sola vez, p.ej. con
`python -c "import secrets; print(secrets.token_hex(32))"`, y no cambiarla
después salvo que se quiera invalidar todas las sesiones recordadas activas.

La planilla debe tener dos pestañas con estos encabezados exactos (fila 1):

    usuarios: usuario | password_hash | nombre | rol | activo
    registro: timestamp | usuario | evento | curso | n_preguntas | n_alumnos |
               dudas_pendientes | sospechosas | sin_identificar | omr_engine_version

Diseño: FALLA CERRADO en el login (cualquier error al leer usuarios o
credenciales inválidas -> sin acceso), pero FALLA ABIERTO en el registro de
actividad (un problema de Sheets -- red, cuota -- nunca debe bloquear el
trabajo real del docente; el error solo se imprime a consola).
"""
import datetime
import sys

import bcrypt
import gspread
from google.oauth2.service_account import Credentials

_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
]

HOJA_USUARIOS = "usuarios"
HOJA_REGISTRO = "registro"
COLUMNAS_USUARIOS = ["usuario", "password_hash", "nombre", "rol", "activo"]
COLUMNAS_REGISTRO = [
    "timestamp", "usuario", "evento", "curso", "n_preguntas", "n_alumnos",
    "dudas_pendientes", "sospechosas", "sin_identificar", "omr_engine_version",
]

_ACTIVO_VALORES_TRUE = {"TRUE", "1", "SI", "SÍ", "X"}


def _secretos():
    """Punto único de acceso a las credenciales -- st.secrets funciona igual
    corriendo dentro de la app (streamlit run) que desde un script suelto
    (gestionar_usuarios.py) siempre que exista .streamlit/secrets.toml en el
    directorio desde donde se ejecuta."""
    import streamlit as st
    return st.secrets


def _cliente_sheets():
    info = dict(_secretos()["gcp_service_account"])
    creds = Credentials.from_service_account_info(info, scopes=_SCOPES)
    return gspread.authorize(creds)


def _spreadsheet():
    return _cliente_sheets().open_by_key(_secretos()["GOOGLE_SHEETS_ID"])


def _hoja(nombre_pestana: str):
    return _spreadsheet().worksheet(nombre_pestana)


def _leer_usuarios() -> list:
    return _hoja(HOJA_USUARIOS).get_all_records()


def _validar_fila_login(fila: dict, password: str) -> dict | None:
    """Lógica PURA de validación sobre una fila ya obtenida de la pestaña
    usuarios -- separada de la lectura real de Sheets para poder probarla
    sin credenciales/red (ver tests/test_registro_sheets.py)."""
    activo = str(fila.get("activo", "")).strip().upper() in _ACTIVO_VALORES_TRUE
    if not activo:
        return None
    hash_guardado = str(fila.get("password_hash", ""))
    try:
        ok = bcrypt.checkpw(password.encode("utf-8"), hash_guardado.encode("utf-8"))
    except (ValueError, TypeError):
        return None
    if not ok:
        return None
    return {
        "usuario": fila.get("usuario"),
        "nombre": fila.get("nombre") or fila.get("usuario"),
        "rol": fila.get("rol") or "docente",
    }


def _buscar_fila_usuario(usuario: str) -> dict | None:
    """Lee la pestaña usuarios y devuelve la fila cruda de ese usuario (o
    None si no existe o hay un error de lectura) -- compartida entre
    verificar_login (exige password) y obtener_usuario_activo (revalida una
    sesión ya recordada, sin password)."""
    try:
        filas = _leer_usuarios()
    except Exception as e:
        print(f"[registro_sheets] error leyendo usuarios: {e}", file=sys.stderr)
        return None
    usuario_norm = usuario.strip().lower()
    for fila in filas:
        if str(fila.get("usuario", "")).strip().lower() == usuario_norm:
            return fila
    return None


def verificar_login(usuario: str, password: str) -> dict | None:
    """Devuelve {"usuario", "nombre", "rol"} si las credenciales son válidas
    y la cuenta está activa, o None en cualquier otro caso -- incluida una
    falla al leer Sheets (fail-closed: sin poder confirmar la credencial,
    no hay acceso)."""
    if not usuario or not password:
        return None
    fila = _buscar_fila_usuario(usuario)
    return _validar_fila_login(fila, password) if fila else None


def obtener_usuario_activo(usuario: str) -> dict | None:
    """Devuelve {"usuario", "nombre", "rol"} si el usuario existe y sigue
    activo, SIN verificar clave -- usado solo para revalidar un token de
    'recordarme' ya firmado (nunca para un login directo, que siempre exige
    password vía verificar_login). Mismo criterio fail-closed: cuenta
    desactivada o error de lectura -> sin acceso, aunque el token en sí
    siga siendo una firma válida."""
    if not usuario:
        return None
    fila = _buscar_fila_usuario(usuario)
    if not fila:
        return None
    activo = str(fila.get("activo", "")).strip().upper() in _ACTIVO_VALORES_TRUE
    if not activo:
        return None
    return {
        "usuario": fila.get("usuario"),
        "nombre": fila.get("nombre") or fila.get("usuario"),
        "rol": fila.get("rol") or "docente",
    }


def registrar_evento(usuario: str, evento: str, **detalle) -> None:
    """Agrega una fila a la pestaña de registro (respaldo de actividad).
    FAIL-OPEN a propósito: nunca debe interrumpir el flujo real de la app
    por un problema de respaldo (red, cuota de la API, permisos).

    value_input_option="RAW": estos son metadatos de auditoría (quién,
    cuándo, curso, contadores), no necesitamos que Sheets interprete nada
    como fórmula -- "curso" en particular viene de un input de texto libre
    del docente, así que con USER_ENTERED un valor que empezara con
    =, +, - o @ se interpretaría como fórmula en la planilla. RAW lo
    guarda siempre como texto literal."""
    try:
        fila = [
            datetime.datetime.now().isoformat(timespec="seconds"),
            usuario,
            evento,
            detalle.get("curso", ""),
            detalle.get("n_preguntas", ""),
            detalle.get("n_alumnos", ""),
            detalle.get("dudas_pendientes", ""),
            detalle.get("sospechosas", ""),
            detalle.get("sin_identificar", ""),
            detalle.get("omr_engine_version", ""),
        ]
        _hoja(HOJA_REGISTRO).append_row(fila, value_input_option="RAW")
    except Exception as e:
        print(f"[registro_sheets] no se pudo registrar evento '{evento}' de "
              f"'{usuario}': {e}", file=sys.stderr)


# ─── Administración de cuentas (usado por gestionar_usuarios.py) ──────────
def agregar_usuario(usuario: str, password: str, nombre: str, rol: str = "docente") -> None:
    """A diferencia de registrar_evento, esto SÍ propaga cualquier error --
    un fallo acá tiene que detener al administrador, no pasar en silencio."""
    password_hash = bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
    # RAW: mismo criterio que registrar_evento -- usuario/nombre son texto
    # simple, no necesitan interpretación de fórmula, y el hash bcrypt debe
    # guardarse literal (RAW no lo altera; ya era compatible con
    # USER_ENTERED porque no empieza con =, +, - ni @).
    _hoja(HOJA_USUARIOS).append_row(
        [usuario, password_hash, nombre, rol, "TRUE"], value_input_option="RAW")


def listar_usuarios() -> list:
    """Nunca incluye el hash de la clave."""
    return [
        {"usuario": f.get("usuario"), "nombre": f.get("nombre"),
         "rol": f.get("rol"), "activo": f.get("activo")}
        for f in _leer_usuarios()
    ]


def desactivar_usuario(usuario: str) -> bool:
    """Pone activo=FALSE en la fila de ese usuario (revoca sin borrar
    historial). Devuelve True si encontró y actualizó la fila."""
    hoja = _hoja(HOJA_USUARIOS)
    usuario_norm = usuario.strip().lower()
    valores_usuario = hoja.col_values(COLUMNAS_USUARIOS.index("usuario") + 1)
    col_activo = COLUMNAS_USUARIOS.index("activo") + 1
    for fila_idx, valor in enumerate(valores_usuario[1:], start=2):  # fila 1 = encabezado
        if str(valor).strip().lower() == usuario_norm:
            hoja.update_cell(fila_idx, col_activo, "FALSE")
            return True
    return False
