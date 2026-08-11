"""
Tests de registro_sheets.py: login (grupo cerrado, cuentas precargadas) y
respaldo de actividad en Google Sheets.

No requieren credenciales ni red -- prueban la lógica PURA
(_validar_fila_login, que ya recibe la fila como dict) y el contrato de
falla de las dos funciones públicas: verificar_login debe FALLAR CERRADO
(sin poder confirmar la credencial, no hay acceso) y registrar_evento debe
FALLAR ABIERTO (un respaldo caído nunca bloquea el trabajo real del
docente) -- ambos simulados con monkeypatch sobre los puntos de acceso a
Sheets (_leer_usuarios, _hoja), sin tocar la red real.

Correr desde la raíz del repo:  py -m pytest tests/test_registro_sheets.py -q
"""
import os
import sys

import bcrypt
import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import registro_sheets as rs


def _fila(password_real="clave123", activo="TRUE", **overrides):
    fila = {
        "usuario": "ana",
        "password_hash": bcrypt.hashpw(password_real.encode("utf-8"), bcrypt.gensalt()).decode("utf-8"),
        "nombre": "Ana Soto",
        "rol": "docente",
        "activo": activo,
    }
    fila.update(overrides)
    return fila


def test_validar_fila_login_credenciales_correctas():
    print("=== _validar_fila_login: usuario activo + clave correcta -> datos de sesion ===")
    resultado = rs._validar_fila_login(_fila(), "clave123")
    assert resultado == {"usuario": "ana", "nombre": "Ana Soto", "rol": "docente"}
    print("OK\n")


def test_validar_fila_login_clave_incorrecta():
    print("=== _validar_fila_login: clave incorrecta -> None ===")
    assert rs._validar_fila_login(_fila(), "clave_mala") is None
    print("OK\n")


def test_validar_fila_login_cuenta_inactiva():
    print("=== _validar_fila_login: cuenta desactivada -> None aunque la clave sea correcta ===")
    assert rs._validar_fila_login(_fila(activo="FALSE"), "clave123") is None
    print("OK\n")


def test_validar_fila_login_hash_corrupto_no_lanza():
    print("=== _validar_fila_login: hash mal formado (fila corrupta en la planilla) -> None, no excepcion ===")
    assert rs._validar_fila_login(_fila(password_hash="esto-no-es-un-hash-bcrypt"), "clave123") is None
    print("OK\n")


def test_verificar_login_usuario_no_existe(monkeypatch):
    print("=== verificar_login: usuario que no esta en la planilla -> None ===")
    monkeypatch.setattr(rs, "_leer_usuarios", lambda: [_fila(usuario="otro")])
    assert rs.verificar_login("ana", "clave123") is None
    print("OK\n")


def test_verificar_login_falla_cerrado_si_sheets_falla(monkeypatch):
    print("=== verificar_login: FALLA CERRADO -- si no se puede leer Sheets, no hay acceso (no excepcion) ===")
    def _explota():
        raise RuntimeError("Sheets no disponible")
    monkeypatch.setattr(rs, "_leer_usuarios", lambda: _explota())
    assert rs.verificar_login("ana", "clave123") is None
    print("OK\n")


def test_verificar_login_credenciales_vacias():
    print("=== verificar_login: usuario o clave vacios -> None sin tocar Sheets ===")
    assert rs.verificar_login("", "clave123") is None
    assert rs.verificar_login("ana", "") is None
    print("OK\n")


class _HojaFalsaQueFalla:
    def append_row(self, *a, **k):
        raise RuntimeError("cuota de la API excedida")


def test_registrar_evento_falla_abierto_si_sheets_falla(monkeypatch):
    print("=== registrar_evento: FALLA ABIERTO -- un respaldo caido NUNCA debe interrumpir el flujo real ===")
    monkeypatch.setattr(rs, "_hoja", lambda nombre: _HojaFalsaQueFalla())
    # No debe lanzar excepcion bajo ninguna circunstancia.
    rs.registrar_evento("ana", "export_excel", curso="3A", n_preguntas=80)
    print("OK -- no se propago la excepcion\n")


class _HojaFalsaQueRegistra:
    def __init__(self):
        self.filas = []

    def append_row(self, fila, value_input_option=None):
        self.filas.append(fila)


def test_registrar_evento_arma_la_fila_en_el_orden_correcto(monkeypatch):
    print("=== registrar_evento: la fila armada respeta el orden de COLUMNAS_REGISTRO ===")
    hoja_falsa = _HojaFalsaQueRegistra()
    monkeypatch.setattr(rs, "_hoja", lambda nombre: hoja_falsa)
    rs.registrar_evento("ana", "export_excel", curso="3A", n_preguntas=80, n_alumnos=25,
                         dudas_pendientes=1, sospechosas=0, sin_identificar=2,
                         omr_engine_version="4.1.1")
    assert len(hoja_falsa.filas) == 1
    fila = hoja_falsa.filas[0]
    assert len(fila) == len(rs.COLUMNAS_REGISTRO)
    idx = {nombre: i for i, nombre in enumerate(rs.COLUMNAS_REGISTRO)}
    assert fila[idx["usuario"]] == "ana"
    assert fila[idx["evento"]] == "export_excel"
    assert fila[idx["curso"]] == "3A"
    assert fila[idx["n_preguntas"]] == 80
    assert fila[idx["n_alumnos"]] == 25
    assert fila[idx["omr_engine_version"]] == "4.1.1"
    print("OK\n")


def test_registrar_evento_login_sin_detalle_no_lanza(monkeypatch):
    print("=== registrar_evento: un evento 'login' (sin curso/preguntas/etc.) no debe fallar por campos faltantes ===")
    hoja_falsa = _HojaFalsaQueRegistra()
    monkeypatch.setattr(rs, "_hoja", lambda nombre: hoja_falsa)
    rs.registrar_evento("ana", "login")
    assert len(hoja_falsa.filas) == 1
    print("OK\n")


if __name__ == "__main__":
    test_validar_fila_login_credenciales_correctas()
    test_validar_fila_login_clave_incorrecta()
    test_validar_fila_login_cuenta_inactiva()
    test_validar_fila_login_hash_corrupto_no_lanza()
    print("TODO PASO - registro_sheets OK (correr con pytest para los tests con monkeypatch)")
