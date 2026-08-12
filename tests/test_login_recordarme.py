"""
Tests del token firmado de "recordarme" (login persistente entre visitas,
ver app_revisor.py: _generar_token_recordar/_validar_token_recordar y el uso
de streamlit_cookies_controller.CookieController alrededor del formulario de
login). El token NUNCA contiene la clave -- es "usuario|expira" firmado con
HMAC-SHA256 usando un secreto de servidor (AUTH_COOKIE_SECRET en st.secrets,
nunca en el repo); revalidar la firma y la expiración es responsabilidad de
estas dos funciones, que reciben el secreto como parámetro explícito (no
llaman a st.secrets) precisamente para poder probarlas sin necesitar
credenciales ni contexto de Streamlit -- mismo criterio que
_validar_fila_login en registro_sheets.py.

Extraídas del código fuente de app_revisor.py con la misma técnica que
tests/test_pauta_excel.py, sin ejecutar la UI.

Correr desde la raíz del repo:  py -m pytest tests/test_login_recordarme.py -q
"""
import base64
import hashlib
import hmac
import os
import re
import sys
import time

import pytest

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, RAIZ)

SECRETO = "secreto-de-pruebas-no-real"
OTRO_SECRETO = "otro-secreto-completamente-distinto"


def _extraer(src, nombre):
    """Igual técnica que test_pauta_excel.py, pero con un límite de fin de
    función más general: acá el código fuente tiene una asignación de nivel
    superior (`_cookies = CookieController(...)`) justo después de la última
    función extraída, no otro `def`/comentario -- el límite original solo
    reconocía esos dos casos y se comía esa línea siguiente."""
    start = src.index(f"\ndef {nombre}") + 1
    rest = src[start + len(f"def {nombre}"):]
    m = re.search(r'\ndef |\n\n\n\S', rest)
    end = start + len(f"def {nombre}") + (m.start() if m else len(rest))
    return src[start:end]


@pytest.fixture(scope="session")
def funcs():
    src = open(os.path.join(RAIZ, "app_revisor.py"), encoding="utf-8").read()
    ns = {"base64": base64, "hashlib": hashlib, "hmac": hmac, "time": time,
          "AUTH_COOKIE_DIAS": 30}
    exec(_extraer(src, "_generar_token_recordar"), ns)
    exec(_extraer(src, "_validar_token_recordar"), ns)
    return ns


def test_token_valido_recupera_el_usuario(funcs):
    print("=== un token recien generado con el secreto correcto se valida y devuelve el usuario ===")
    token = funcs["_generar_token_recordar"]("ana", SECRETO)
    assert token is not None
    assert funcs["_validar_token_recordar"](token, SECRETO) == "ana"
    print("OK\n")


def test_sin_secreto_no_genera_ni_valida_nada(funcs):
    print("=== fail-closed: sin AUTH_COOKIE_SECRET configurado, no se genera ni se valida ningun token ===")
    assert funcs["_generar_token_recordar"]("ana", "") is None
    token = funcs["_generar_token_recordar"]("ana", SECRETO)
    assert funcs["_validar_token_recordar"](token, "") is None
    print("OK\n")


def test_token_firmado_con_otro_secreto_no_se_valida(funcs):
    print("=== un token firmado con un secreto distinto (rotado, o falsificado) debe rechazarse ===")
    token = funcs["_generar_token_recordar"]("ana", OTRO_SECRETO)
    assert funcs["_validar_token_recordar"](token, SECRETO) is None
    print("OK\n")


def test_token_expirado_no_se_valida(funcs):
    print("=== un token cuya fecha de expiracion ya paso debe rechazarse, aunque la firma sea correcta ===")
    token = funcs["_generar_token_recordar"]("ana", SECRETO, dias=-1)  # "expiro" hace un dia
    assert funcs["_validar_token_recordar"](token, SECRETO) is None
    print("OK\n")


def test_token_manipulado_no_se_valida(funcs):
    print("=== cambiar el usuario dentro de un token ajeno (sin conocer el secreto) debe fallar la firma ===")
    token_ana = funcs["_generar_token_recordar"]("ana", SECRETO)
    token_admin = funcs["_generar_token_recordar"]("admin", SECRETO)
    payload_admin = token_admin.split(".")[0]
    firma_ana = token_ana.split(".")[1]
    token_falsificado = f"{payload_admin}.{firma_ana}"  # payload de "admin" + firma real de "ana"
    assert funcs["_validar_token_recordar"](token_falsificado, SECRETO) is None
    print("OK\n")


def test_valores_invalidos_no_lanzan_excepcion(funcs):
    print("=== texto arbitrario, vacio o None como token nunca debe lanzar excepcion ===")
    assert funcs["_validar_token_recordar"]("", SECRETO) is None
    assert funcs["_validar_token_recordar"](None, SECRETO) is None
    assert funcs["_validar_token_recordar"]("esto-no-es-un-token-valido", SECRETO) is None
    assert funcs["_validar_token_recordar"]("sin-punto-de-separacion", SECRETO) is None
    print("OK\n")


def test_usuario_con_caracteres_especiales_sobrevive_el_roundtrip(funcs):
    print("=== un usuario con '|' o '.' en el nombre no debe romper el parseo del payload ===")
    token = funcs["_generar_token_recordar"]("ana|con.puntos|raros", SECRETO)
    assert funcs["_validar_token_recordar"](token, SECRETO) == "ana|con.puntos|raros"
    print("OK\n")


def test_usuario_vacio_no_genera_token(funcs):
    print("=== sin usuario no hay nada que recordar ===")
    assert funcs["_generar_token_recordar"]("", SECRETO) is None
    assert funcs["_generar_token_recordar"](None, SECRETO) is None
    print("OK\n")


if __name__ == "__main__":
    src = open(os.path.join(RAIZ, "app_revisor.py"), encoding="utf-8").read()
    ns = {"base64": base64, "hashlib": hashlib, "hmac": hmac, "time": time, "AUTH_COOKIE_DIAS": 30}
    exec(_extraer(src, "_generar_token_recordar"), ns)
    exec(_extraer(src, "_validar_token_recordar"), ns)
    test_token_valido_recupera_el_usuario(ns)
    test_sin_secreto_no_genera_ni_valida_nada(ns)
    test_token_firmado_con_otro_secreto_no_se_valida(ns)
    test_token_expirado_no_se_valida(ns)
    test_token_manipulado_no_se_valida(ns)
    test_valores_invalidos_no_lanzan_excepcion(ns)
    test_usuario_con_caracteres_especiales_sobrevive_el_roundtrip(ns)
    test_usuario_vacio_no_genera_token(ns)
    print("TODO PASO - token de recordarme OK")
