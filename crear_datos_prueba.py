# -*- coding: utf-8 -*-
# ============================================================================
# crear_datos_prueba.py — Generador de datos y trabajos de prueba para V2
# ============================================================================
"""
Crea archivos de prueba en TRABAJOS_ESTUDIANTES/ para verificar la evaluación
en Modo A (Plantilla Espejo) y Modo B (Rúbrica Inteligente).
"""

import os
import sys
import shutil
import openpyxl

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

from core.config import PLANTILLAS_DIR, TRABAJOS_DIR, DEFAULT_PLANTILLA_PATH
from core.rubric_template import generar_plantilla_rubrica


def crear_escenarios_de_prueba():
    os.makedirs(PLANTILLAS_DIR, exist_ok=True)
    os.makedirs(TRABAJOS_DIR, exist_ok=True)

    # 1. Crear plantilla base
    ruta_plantilla_ejemplo = os.path.join(PLANTILLAS_DIR, "PLANTILLA_RUBRICA_EJEMPLO.xlsx")
    generar_plantilla_rubrica(ruta_plantilla_ejemplo)

    # También copiar como PLANTILLA.xlsx si no existe
    if not os.path.exists(DEFAULT_PLANTILLA_PATH):
        shutil.copy(ruta_plantilla_ejemplo, DEFAULT_PLANTILLA_PATH)

    print(f"✅ Plantilla modelo creada en: {ruta_plantilla_ejemplo}")

    # 2. Estudiante 1: Trabajo Perfecto (debe obtener calificación alta/aprobado)
    est1_path = os.path.join(TRABAJOS_DIR, "Estudiante_1_Excelente.xlsx")
    shutil.copy(ruta_plantilla_ejemplo, est1_path)
    print(f"✅ Creado: {os.path.basename(est1_path)}")

    # 3. Estudiante 2: Trabajo con Errores Puntuales
    est2_path = os.path.join(TRABAJOS_DIR, "Estudiante_2_Parcial.xlsx")
    wb2 = openpyxl.load_workbook(ruta_plantilla_ejemplo)
    ws2_v = wb2["Ventas"]
    # Error en G11: valor fijo en lugar de SUMA
    ws2_v["G11"] = 1200.0
    # Error en formato: quitar formato de moneda en D2:D5
    for r in range(2, 6):
        ws2_v[f"D{r}"].number_format = "General"
    # Error en negrita de encabezados
    ws2_v["A1"].font = openpyxl.styles.Font(bold=False)

    wb2.save(est2_path)
    wb2.close()
    print(f"✅ Creado: {os.path.basename(est2_path)}")

    # 4. Estudiante 3: Trabajo Desplazado (Insertó una fila en blanco al inicio)
    est3_path = os.path.join(TRABAJOS_DIR, "Estudiante_3_Desplazado.xlsx")
    wb3 = openpyxl.load_workbook(ruta_plantilla_ejemplo)
    ws3_v = wb3["Ventas"]
    # Insertar fila vacía en la fila 1
    ws3_v.insert_rows(1)
    wb3.save(est3_path)
    wb3.close()
    print(f"✅ Creado: {os.path.basename(est3_path)}")


if __name__ == "__main__":
    crear_escenarios_de_prueba()
