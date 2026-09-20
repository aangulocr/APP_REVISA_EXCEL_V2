# -*- coding: utf-8 -*-
# ============================================================================
# main.py — Punto de entrada CLI para APP_REVISA_EXCEL_V2
# ============================================================================
"""
Interfaz de línea de comandos (CLI) para ejecutar la auditoría híbrida
de archivos de Excel.

Uso:
    python main.py
    python main.py --modo modo_b_rubrica --plantilla PLANTILLAS/PLANTILLA.xlsx
    python main.py --generar-plantilla
"""

import os
import sys
import argparse
import logging
from datetime import datetime

# Configurar salida UTF-8 segura en Windows para emojis y caracteres especiales
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

from core.config import (
    DEFAULT_PLANTILLA_PATH,
    TRABAJOS_DIR,
    REVISADOS_DIR,
    LOGS_DIR,
)
from core.models import ModoEvaluacion
from core.rubric_template import generar_plantilla_rubrica
from auditor import auditar_lote


def configurar_logging(carpeta_logs: str = LOGS_DIR, seccion: str = "AUDITORIA"):
    """Configura logging para consola y archivo de registro."""
    os.makedirs(carpeta_logs, exist_ok=True)
    fecha_str = datetime.now().strftime("%d%m%y_%H%M%S")
    log_file = os.path.join(carpeta_logs, f"{seccion}_{fecha_str}.log")

    log_format = "%(asctime)s | %(levelname)-7s | %(message)s"
    date_format = "%Y-%m-%d %H:%M:%S"

    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    logger.handlers = []

    # Consola
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    ch.setFormatter(logging.Formatter(log_format, datefmt=date_format))
    logger.addHandler(ch)

    # Archivo
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter(log_format, datefmt=date_format))
    logger.addHandler(fh)

    return log_file


def main():
    parser = argparse.ArgumentParser(
        description="APP_REVISA_EXCEL_V2 — Motor Híbrido de Auditoría y Evaluación de Excel"
    )
    parser.add_argument(
        "--plantilla",
        type=str,
        default=DEFAULT_PLANTILLA_PATH,
        help=f"Ruta a la plantilla de Excel (.xlsx). Por defecto: {DEFAULT_PLANTILLA_PATH}"
    )
    parser.add_argument(
        "--trabajos",
        type=str,
        default=TRABAJOS_DIR,
        help=f"Carpeta con los trabajos de estudiantes. Por defecto: {TRABAJOS_DIR}"
    )
    parser.add_argument(
        "--modo",
        type=str,
        choices=["auto", "modo_a_espejo", "modo_b_rubrica", "espejo", "rubrica"],
        default="auto",
        help="Modo de evaluación: 'auto' (detecta si hay rúbrica), 'modo_a_espejo' o 'modo_b_rubrica'."
    )
    parser.add_argument(
        "--seccion",
        type=str,
        default="SEC_DEFAULT",
        help="Nombre de la sección o grupo (ej: 11-2B)"
    )
    parser.add_argument(
        "--fecha",
        type=str,
        default="",
        help="Fecha opcional para el reporte (formato libre o ddMMYY)"
    )
    parser.add_argument(
        "--generar-plantilla",
        action="store_true",
        help="Genera un archivo de plantilla con hoja _RUBRICA prellenada a 100 puntos y finaliza."
    )

    args = parser.parse_args()

    # Si se solicitó solo generar la plantilla de rúbrica
    if args.generar_plantilla:
        print("Generando plantilla de ejemplo con rúbrica...")
        ruta_gen = generar_plantilla_rubrica()
        print(f"✅ Plantilla de ejemplo creada exitosamente en: {ruta_gen}")
        sys.exit(0)

    # Si la plantilla por defecto no existe pero existe la de ejemplo, usarla
    ruta_plantilla_final = args.plantilla
    if not os.path.exists(ruta_plantilla_final):
        plantilla_ejemplo = os.path.join(os.path.dirname(ruta_plantilla_final), "PLANTILLA_RUBRICA_EJEMPLO.xlsx")
        if os.path.exists(plantilla_ejemplo):
            ruta_plantilla_final = plantilla_ejemplo

    # Resolver modo
    modo_sel = None
    if args.modo != "auto":
        modo_sel = ModoEvaluacion.from_str(args.modo)

    # Iniciar logging
    log_file = configurar_logging(seccion=args.seccion)

    print("=" * 65)
    print("      APP_REVISA_EXCEL_V2 — MOTOR HÍBRIDO DE EVALUACIÓN")
    print("=" * 65)
    print(f"Modo seleccionado : {args.modo.upper()}")
    print(f"Plantilla         : {ruta_plantilla_final}")
    print(f"Carpeta trabajos  : {args.trabajos}")
    print(f"Archivo de log    : {log_file}")
    print("=" * 65)

    try:
        resultados = auditar_lote(
            ruta_plantilla=ruta_plantilla_final,
            ruta_trabajos=args.trabajos,
            modo=modo_sel,
            seccion=args.seccion
        )
        print("\n" + "=" * 65)
        print(f"✅ Proceso finalizado. Se evaluaron {len(resultados)} archivos.")
        print(f"Archivos revisados guardados en: {REVISADOS_DIR}")
        print(f"Reportes guardados en          : {LOGS_DIR}")
        print("=" * 65)
        sys.exit(0)
    except Exception as e:
        logging.error(f"Error durante la auditoría: {e}", exc_info=True)
        print(f"\n❌ Error durante la auditoría: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
