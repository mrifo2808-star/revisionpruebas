"""
Tests de "Importar rápido" (pegar la pauta como texto libre en la pestaña
Pauta de respuestas, ver app_revisor.py: extraer_letras_pauta_rapida).

Antes se usaba re.findall(r'[AaBbCcDdEe]', texto), que interpretaba
cualquier letra A-E suelta dentro de OTRAS palabras (p.ej. la "a" de
"Alternativa"). extraer_letras_pauta_rapida exige que la letra esté aislada
por un límite de palabra (\\b) -- separada por espacio, salto de línea,
coma, paréntesis, "=", o inicio/fin de texto -- para contarla como una
respuesta real.

Extraída del código fuente de app_revisor.py con la misma técnica que
tests/test_pauta_excel.py, sin ejecutar la UI.

Correr desde la raíz del repo:  py -m pytest tests/test_pauta_importar_rapido.py -q
"""
import os
import re
import sys

import pytest

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, RAIZ)


def _extraer(src, nombre):
    start = src.index(f"\ndef {nombre}") + 1
    rest = src[start + len(f"def {nombre}"):]
    m = re.search(r'\ndef |\n\n\n\S', rest)
    end = start + len(f"def {nombre}") + (m.start() if m else len(rest))
    return src[start:end]


@pytest.fixture(scope="session")
def extraer_letras_pauta_rapida():
    src = open(os.path.join(RAIZ, "app_revisor.py"), encoding="utf-8").read()
    ns = {"re": re}
    exec(_extraer(src, "extraer_letras_pauta_rapida"), ns)
    return ns["extraer_letras_pauta_rapida"]


# ─── Casos positivos: formatos reales que SÍ deben reconocerse ────────────

@pytest.mark.parametrize("texto,esperado", [
    ("A\nB\nC\nD", ["A", "B", "C", "D"]),
    ("A,B,C,D", ["A", "B", "C", "D"]),
    ("A B C D", ["A", "B", "C", "D"]),
    ("a b c d", ["A", "B", "C", "D"]),
    ("1 A\n2 B\n3 C", ["A", "B", "C"]),
    ("P1=A\nP2=B", ["A", "B"]),
    ("(A) (B) (C)", ["A", "B", "C"]),
    ("A", ["A"]),
])
def test_formatos_reales_se_reconocen(extraer_letras_pauta_rapida, texto, esperado):
    print(f"=== formato real '{texto!r}' debe producir {esperado} ===")
    letras = [l.upper() for l in extraer_letras_pauta_rapida(texto)]
    assert letras == esperado
    print("OK\n")


# ─── Casos negativos: texto narrativo que NO debe inventar una pauta ──────

@pytest.mark.parametrize("texto", [
    "Alternativa",
    "Respuesta",
    "ABCDEtexto",
    "Prueba de lenguaje",
    "",
    "   ",
    "Cadena",
    "Decisión",
])
def test_texto_narrativo_no_inventa_pauta(extraer_letras_pauta_rapida, texto):
    print(f"=== texto narrativo '{texto!r}' no debe extraer ninguna letra ===")
    letras = extraer_letras_pauta_rapida(texto)
    assert letras == []
    print("OK\n")


def test_letra_dentro_de_palabra_no_se_cuenta(extraer_letras_pauta_rapida):
    print("=== 'Alternativa A' solo cuenta la A aislada al final, no las 'a' internas ===")
    letras = [l.upper() for l in extraer_letras_pauta_rapida("Alternativa A")]
    assert letras == ["A"]
    print("OK\n")


def test_none_no_lanza_excepcion(extraer_letras_pauta_rapida):
    print("=== texto None no debe lanzar excepcion ===")
    assert extraer_letras_pauta_rapida(None) == []
    print("OK\n")
