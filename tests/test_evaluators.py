# -*- coding: utf-8 -*-
"""
Pruebas de integración de evaluadores (Modo A y Modo B) y reportes.
"""

import os
import shutil
import tempfile
import pytest
import openpyxl

from core.rubric_template import generar_plantilla_rubrica
from core.evaluators import RubricEvaluator, MirrorEvaluator
from core.report_generator import (
    exportar_csv,
    exportar_json_detallado,
    exportar_excel_log,
    exportar_observaciones_txt,
)


@pytest.fixture
def temp_environment():
    temp_dir = tempfile.mkdtemp()
    plantilla_path = os.path.join(temp_dir, "PLANTILLA.xlsx")
    generar_plantilla_rubrica(plantilla_path)

    trabajos_dir = os.path.join(temp_dir, "TRABAJOS")
    revisados_dir = os.path.join(temp_dir, "REVISADOS")
    logs_dir = os.path.join(temp_dir, "LOGS")
    os.makedirs(trabajos_dir, exist_ok=True)
    os.makedirs(revisados_dir, exist_ok=True)
    os.makedirs(logs_dir, exist_ok=True)

    # Crear estudiante 1: copia idéntica (debe aprobar 100)
    est1_path = os.path.join(trabajos_dir, "Estudiante_Excelente.xlsx")
    shutil.copy(plantilla_path, est1_path)

    # Crear estudiante 2: con errores puntuales
    est2_path = os.path.join(trabajos_dir, "Estudiante_Errores.xlsx")
    wb_err = openpyxl.load_workbook(plantilla_path)
    # Borrar fórmula en G11 y poner número fijo
    ws_v = wb_err["Ventas"]
    ws_v["G11"] = 100.0  # Erróneo
    wb_err.save(est2_path)
    wb_err.close()

    yield {
        "temp_dir": temp_dir,
        "plantilla": plantilla_path,
        "trabajos": trabajos_dir,
        "revisados": revisados_dir,
        "logs": logs_dir,
        "est1": est1_path,
        "est2": est2_path,
    }

    shutil.rmtree(temp_dir, ignore_errors=True)


def test_rubric_evaluator_integration(temp_environment):
    env = temp_environment
    evaluador = RubricEvaluator()

    # Evaluar estudiante 1
    res1 = evaluador.evaluar(env["plantilla"], env["est1"], env["revisados"])
    assert res1.puntos_obtenidos >= 90.0
    assert res1.aprobado is True
    assert os.path.isfile(res1.archivo_revisado)

    # Evaluar estudiante 2
    res2 = evaluador.evaluar(env["plantilla"], env["est2"], env["revisados"])
    assert res2.puntos_obtenidos < res1.puntos_obtenidos
    assert len(res2.criterios_fallados) > 0

    # Comprobar que en el archivo revisado se insertaron comentarios de error
    wb_rev = openpyxl.load_workbook(res2.archivo_revisado)
    ws_rev = wb_rev["Ventas"]
    celda_g11 = ws_rev["G11"]
    assert celda_g11.comment is not None
    assert "FALLÓ" in celda_g11.comment.text
    wb_rev.close()


def test_report_generation(temp_environment):
    env = temp_environment
    evaluador = RubricEvaluator()
    res1 = evaluador.evaluar(env["plantilla"], env["est1"], env["revisados"])
    res2 = evaluador.evaluar(env["plantilla"], env["est2"], env["revisados"])
    resultados = [res1, res2]

    # Exportar CSV
    csv_path = os.path.join(env["logs"], "REPORTE_NOTAS.csv")
    exportar_csv(resultados, csv_path)
    assert os.path.isfile(csv_path)

    # Exportar JSON
    json_path = os.path.join(env["logs"], "REPORTE_DETALLADO.json")
    exportar_json_detallado(resultados, json_path)
    assert os.path.isfile(json_path)

    # Exportar Excel
    xlsx_path = os.path.join(env["logs"], "LOG_NOTAS.xlsx")
    exportar_excel_log(resultados, xlsx_path)
    assert os.path.isfile(xlsx_path)

    # Exportar Observaciones en texto plano
    txt_obs_path = os.path.join(env["logs"], "OBSERVACIONES_ESTUDIANTES.txt")
    exportar_observaciones_txt(resultados, txt_obs_path, seccion="11-1")
    assert os.path.isfile(txt_obs_path)
    with open(txt_obs_path, "r", encoding="utf-8") as f:
        contenido_obs = f.read()
        assert "OBSERVACIONES Y FALLOS DETECTADOS" in contenido_obs

    # Verificar contenido de Excel
    wb_log = openpyxl.load_workbook(xlsx_path)
    assert "Resumen General" in wb_log.sheetnames
    assert "Matriz de Criterios" in wb_log.sheetnames
    ws_res = wb_log["Resumen General"]
    # Columna 9 debe ser la de Observaciones
    header_col9 = ws_res.cell(row=1, column=9).value
    assert "OBSERVACIONES" in str(header_col9).upper()
    wb_log.close()
