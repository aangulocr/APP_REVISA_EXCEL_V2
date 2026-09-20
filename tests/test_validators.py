# -*- coding: utf-8 -*-
"""
Pruebas unitarias para el catálogo de validadores del Modo B.
"""

import pytest
import openpyxl
from core.models import CriterioRubrica
from core.validators.value_validator import ValueValidator
from core.validators.formula_validator import FormulaValidator
from core.validators.format_validator import FormatValidator
from core.validators.style_validator import StyleValidator
from core.validators.structure_validator import StructureValidator


def test_value_validator_numeric():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = 1500.04

    val = ValueValidator()

    # Coincidencia con tolerancia 0.05
    crit_ok = CriterioRubrica(
        id="C1",
        criterio="Total 1500",
        hoja="Sheet",
        rango="A1",
        tipo_validacion="valor_numerico",
        parametros={"valor": 1500.0, "tolerancia": 0.05},
        puntos=10
    )
    res_ok = val.validar(ws, crit_ok)
    assert res_ok.aprobado is True
    assert res_ok.puntos_obtenidos == 10

    # Falla con tolerancia estricta
    crit_fail = CriterioRubrica(
        id="C2",
        criterio="Total 1500 estricto",
        hoja="Sheet",
        rango="A1",
        tipo_validacion="valor_numerico",
        parametros={"valor": 1500.0, "tolerancia": 0.01},
        puntos=10
    )
    res_fail = val.validar(ws, crit_fail)
    assert res_fail.aprobado is False
    assert res_fail.puntos_obtenidos == 0


def test_value_validator_text_normalized():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["B2"] = "  MÉXICO  "

    val = ValueValidator()
    crit = CriterioRubrica(
        id="C3",
        criterio="País México",
        hoja="Sheet",
        rango="B2",
        tipo_validacion="valor_exacto",
        parametros={"valor": "mexico"},
        puntos=15
    )
    res = val.validar(ws, crit)
    assert res.aprobado is True
    assert res.puntos_obtenidos == 15


def test_formula_validator_bilingual_and_refs():
    wb = openpyxl.Workbook()
    ws = wb.active
    # En openpyxl las fórmulas pueden estar escritas como SUM(C2:C10) o =SUMA(C2:C10)
    ws["D20"] = "=SUMA(C2:C10)"

    val = FormulaValidator()

    # Validar que reconoce SUM o SUMA
    crit = CriterioRubrica(
        id="C4",
        criterio="Suma de columna C",
        hoja="Sheet",
        rango="D20",
        tipo_validacion="formula_flexible",
        parametros={
            "funciones": ["SUM"],
            "referencias": ["C2:C10"]
        },
        puntos=20
    )
    res = val.validar(ws, crit)
    assert res.aprobado is True
    assert res.puntos_obtenidos == 20


def test_format_validator():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["E5"] = 1250.5
    ws["E5"].number_format = '$#,##0.00'

    ws["F5"] = 0.15
    ws["F5"].number_format = '0.00%'

    val = FormatValidator()

    # Moneda
    crit_moneda = CriterioRubrica(
        id="C5",
        criterio="Moneda en E5",
        hoja="Sheet",
        rango="E5",
        tipo_validacion="formato_numero",
        parametros={"formato": "moneda"},
        puntos=10
    )
    assert val.validar(ws, crit_moneda).aprobado is True

    # Porcentaje
    crit_pct = CriterioRubrica(
        id="C6",
        criterio="Porcentaje en F5",
        hoja="Sheet",
        rango="F5",
        tipo_validacion="formato_numero",
        parametros={"formato": "porcentaje"},
        puntos=10
    )
    assert val.validar(ws, crit_pct).aprobado is True


def test_style_validator():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "Encabezado"
    ws["A1"].font = openpyxl.styles.Font(bold=True)
    ws["A1"].fill = openpyxl.styles.PatternFill(start_color="FF0056B3", end_color="FF0056B3", fill_type="solid")

    val = StyleValidator()
    crit = CriterioRubrica(
        id="C7",
        criterio="Encabezado negrita y relleno",
        hoja="Sheet",
        rango="A1",
        tipo_validacion="estilo_visual_flexible",
        parametros={"requiere_negrita": True, "requiere_relleno": True},
        puntos=10
    )
    assert val.validar(ws, crit).aprobado is True


def test_structure_validator_merged():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.merge_cells("A1:D1")

    val = StructureValidator()
    crit = CriterioRubrica(
        id="C8",
        criterio="Título combinado A1:D1",
        hoja="Sheet",
        rango="A1:D1",
        tipo_validacion="celdas_combinadas",
        parametros={},
        puntos=10
    )
    assert val.validar(ws, crit).aprobado is True


def test_chart_validator():
    from core.validators.chart_validator import ChartValidator
    wb = openpyxl.Workbook()
    cs = wb.create_chartsheet(title="Gráfico Ventas x Vendedor")

    val = ChartValidator()
    crit_chartsheet = CriterioRubrica(
        id="C9",
        criterio="Mover gráfico a hoja nueva Gráfico Ventas x Vendedor",
        hoja="Gráfico Ventas x Vendedor",
        rango="A1",
        tipo_validacion="grafico_dinamico",
        parametros={"tipo_hoja": "chartsheet", "origen": "Resumen_Ventas"},
        puntos=20
    )
    res = val.validar(cs, crit_chartsheet)
    assert res.aprobado is True
    assert res.puntos_obtenidos == 20.0
    assert "ChartSheet" in res.mensaje_detalle

    # Si se pasa una hoja normal cuando se requería chartsheet
    ws_normal = wb.active
    res_err = val.validar(ws_normal, crit_chartsheet)
    assert res_err.aprobado is False
    assert res_err.puntos_obtenidos == 0.0


def test_combined_operations_and_ifs():
    from core.validators.formula_validator import FormulaValidator
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["F3"] = '=SI.CONJUNTO(D3="JEFE", C3*E3 - C3*E3*5%, D3="VENDEDOR", C3*E3 - C3*E3*2%)'
    ws["G3"] = '= D3 * E3 - D3 * E3 * 2%'

    val = FormulaValidator()

    # Caso 1: SI.CONJUNTO con operacion combinada adentro
    crit1 = CriterioRubrica(
        id="C10",
        criterio="SI.CONJUNTO con descuento",
        hoja="Sheet",
        rango="F3",
        tipo_validacion="formula",
        parametros={"funciones": ["SI.CONJUNTO"], "referencias": ["D3", "C3", "E3"], "contiene": ["JEFE", "5%/0.05"]},
        puntos=15
    )
    res1 = val.validar(ws, crit1)
    assert res1.aprobado is True
    assert res1.puntos_obtenidos == 15.0

    # Caso 2: Operacion combinada directa
    crit2 = CriterioRubrica(
        id="C11",
        criterio="Operación combinada de descuento",
        hoja="Sheet",
        rango="G3",
        tipo_validacion="formula",
        parametros={"referencias": ["D3", "E3"], "contiene": ["2%/0.02"]},
        puntos=10
    )
    res2 = val.validar(ws, crit2)
    assert res2.aprobado is True
    assert res2.puntos_obtenidos == 10.0


