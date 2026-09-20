# -*- coding: utf-8 -*-
"""
Pruebas unitarias para el parser de rúbrica (Excel y JSON) y generador de plantilla.
"""

import os
import tempfile
import json
import pytest
import openpyxl

from core.rubric_parser import (
    parsear_rubrica_desde_hoja,
    parsear_rubrica_desde_json,
    cargar_rubrica,
)
from core.rubric_template import generar_plantilla_rubrica


def test_generar_y_parsear_plantilla_rubrica():
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp_path = tmp.name

    try:
        # Generar plantilla de ejemplo
        generar_plantilla_rubrica(tmp_path)
        assert os.path.isfile(tmp_path)

        # Cargar rúbrica
        criterios, puntos_totales, advertencias, origen = cargar_rubrica(tmp_path)

        assert len(criterios) == 8
        assert abs(puntos_totales - 100.0) < 0.01
        assert "Hoja Excel" in origen
        assert any(c.id == "CRIT_01" for c in criterios)
        assert any(c.tipo_validacion == "formula_flexible" for c in criterios)
        assert any(c.tipo_validacion == "formato_numero" for c in criterios)
        assert any(c.tipo_validacion == "celdas_combinadas" for c in criterios)
    finally:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass


def test_parsear_rubrica_desde_json():
    datos_json = {
        "criterios": [
            {
                "id": "CRIT_01",
                "criterio": "Total con SUMA",
                "hoja": "Ventas",
                "rango": "D20",
                "tipo_validacion": "formula_flexible",
                "parametros": {"funciones": ["SUM"]},
                "puntos": 50,
                "obligatorio": True
            },
            {
                "id": "CRIT_02",
                "criterio": "Valor exacto",
                "hoja": "Ventas",
                "rango": "D20",
                "tipo_validacion": "valor_numerico",
                "parametros": {"valor": 1000},
                "puntos": 50,
                "obligatorio": False
            }
        ]
    }

    with tempfile.NamedTemporaryFile(suffix=".json", mode="w", encoding="utf-8", delete=False) as tmp:
        json.dump(datos_json, tmp)
        tmp_path = tmp.name

    try:
        criterios, pts, advs = parsear_rubrica_desde_json(tmp_path)
        assert len(criterios) == 2
        assert pts == 100.0
        assert criterios[0].obligatorio is True
        assert criterios[1].obligatorio is False
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
