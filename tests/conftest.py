"""
Fixture compartido para que los tests de tests/*.py corran bajo pytest
(`python -m pytest tests -q`) además de como script standalone
(`py tests/test_xxx.py`). Ambos test_omr_regression.py y
test_omr_invariantes.py escriben sus funciones como `def test_xxx(app):` --
pytest resuelve ese parámetro contra esta fixture automáticamente por nombre.
"""
import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import omr_metrics


@pytest.fixture(scope="session")
def app():
    return omr_metrics._cargar_app_module()
