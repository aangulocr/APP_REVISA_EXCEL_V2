# -*- coding: utf-8 -*-
# ============================================================================
# auditor.py — Fachada orquestadora principal de APP_REVISA_EXCEL_V2
# ============================================================================
"""
Coordina la ejecución completa del proceso de auditoría:
  - Detección o selección del Modo de Evaluación (A: Espejo / B: Rúbrica)
  - Iteración sobre el lote de archivos de estudiantes
  - Ejecución de evaluadores con retroalimentación en tiempo real
  - Generación de reportes consolidados (CSV, JSON, Excel)
"""

import os
import glob
import logging
from datetime import datetime
from typing import List, Optional, Callable, Dict, Any

from core.config import (
    DEFAULT_PLANTILLA_PATH,
    TRABAJOS_DIR,
    REVISADOS_DIR,
    LOGS_DIR,
    EXTENSIONES_VALIDAS,
)
from core.models import ModoEvaluacion, ResultadoEstudiante
from core.rubric_parser import cargar_rubrica
from core.evaluators import MirrorEvaluator, RubricEvaluator
from core.report_generator import (
    exportar_csv,
    exportar_json_detallado,
    exportar_excel_log,
    exportar_observaciones_txt,
)

logger = logging.getLogger(__name__)


def detectar_modo_automatico(ruta_plantilla: str) -> ModoEvaluacion:
    """
    Detecta automáticamente si la plantilla cuenta con una rúbrica configurada
    (en hoja _RUBRICA o archivo .json). Si existe, usa Modo B; si no, Modo A.
    """
    criterios, _, _, origen = cargar_rubrica(ruta_plantilla)
    if criterios:
        logger.info(f"Rúbrica detectada automáticamente desde: {origen} -> Activando Modo B.")
        return ModoEvaluacion.MODO_B_RUBRICA
    logger.info("No se detectó rúbrica -> Activando Modo A (Plantilla Espejo).")
    return ModoEvaluacion.MODO_A_ESPEJO


def auditar_lote(
    ruta_plantilla: str,
    ruta_trabajos: str,
    modo: Optional[ModoEvaluacion] = None,
    carpeta_revisados: str = REVISADOS_DIR,
    carpeta_logs: str = LOGS_DIR,
    seccion: str = "SEC_DEFAULT",
    callback_progreso: Optional[Callable[[Dict[str, Any]], None]] = None
) -> List[ResultadoEstudiante]:
    """
    Ejecuta la auditoría sobre todos los archivos de estudiantes encontrados.

    Parámetros:
        ruta_plantilla: Ruta al archivo PLANTILLA.xlsx.
        ruta_trabajos: Carpeta con los archivos de los estudiantes.
        modo: ModoEvaluacion (MODO_A_ESPEJO o MODO_B_RUBRICA). Si es None, se autodetecta.
        carpeta_revisados: Directorio donde se guardarán los archivos calificados (_rev.xlsx).
        carpeta_logs: Directorio para guardar reportes y logs.
        seccion: Identificador del grupo o sección escolar.
        callback_progreso: Función opcional para notificar avances (porcentaje, archivo actual).

    Retorna:
        Lista de ResultadoEstudiante.
    """
    if not os.path.isfile(ruta_plantilla):
        raise FileNotFoundError(f"No se encontró el archivo de plantilla: '{ruta_plantilla}'")

    if not os.path.isdir(ruta_trabajos):
        raise FileNotFoundError(f"No se encontró la carpeta de trabajos: '{ruta_trabajos}'")

    # Identificar archivos válidos
    archivos = []
    for f in sorted(os.listdir(ruta_trabajos)):
        if f.lower().endswith(EXTENSIONES_VALIDAS) and not f.startswith("~$") and not f.startswith("."):
            archivos.append(os.path.join(ruta_trabajos, f))

    if not archivos:
        raise ValueError(f"No se encontraron archivos .xlsx válidos en '{ruta_trabajos}'")

    # Determinar modo si no fue provisto
    if modo is None:
        modo = detectar_modo_automatico(ruta_plantilla)

    logger.info(f"=== INICIANDO AUDITORÍA EN {modo.value.upper()} ===")
    logger.info(f"Plantilla: {ruta_plantilla}")
    logger.info(f"Archivos de estudiantes a evaluar: {len(archivos)}")

    # Instanciar evaluador
    if modo == ModoEvaluacion.MODO_B_RUBRICA:
        criterios, pts_totales, advs, origen = cargar_rubrica(ruta_plantilla)
        evaluador = RubricEvaluator(criterios_predefinidos=criterios)
        logger.info(f"Rúbrica cargada desde {origen} ({len(criterios)} criterios, {pts_totales} pts).")
    else:
        evaluador = MirrorEvaluator()

    resultados: List[ResultadoEstudiante] = []
    total_archivos = len(archivos)

    for idx, ruta_est in enumerate(archivos, 1):
        nombre_base = os.path.basename(ruta_est)
        logger.info(f"[{idx}/{total_archivos}] Evaluando '{nombre_base}'...")

        if callback_progreso:
            callback_progreso({
                "evento": "inicio_archivo",
                "archivo": nombre_base,
                "indice": idx,
                "total": total_archivos,
                "porcentaje": round(((idx - 1) / total_archivos) * 100, 1)
            })

        res = evaluador.evaluar(
            ruta_plantilla=ruta_plantilla,
            ruta_estudiante=ruta_est,
            carpeta_salida_revisados=carpeta_revisados
        )
        resultados.append(res)

        estado_txt = "APROBADO" if res.aprobado else "REPROBADO"
        logger.info(
            f"    -> Nota: {res.nota_final_100:.1f}/100 [{estado_txt}] "
            f"({res.tiempo_segundos:.2f}s)"
        )

        if callback_progreso:
            callback_progreso({
                "evento": "fin_archivo",
                "archivo": nombre_base,
                "indice": idx,
                "total": total_archivos,
                "porcentaje": round((idx / total_archivos) * 100, 1),
                "resultado": res.to_dict()
            })

    # Generar reportes consolidados
    os.makedirs(carpeta_logs, exist_ok=True)
    fecha_str = datetime.now().strftime("%d%m%y_%H%M%S")

    # CSV
    ruta_csv = os.path.join(carpeta_logs, "REPORTE_NOTAS.csv")
    exportar_csv(resultados, ruta_csv)
    logger.info(f"Reporte CSV generado en: {ruta_csv}")

    # JSON Detallado
    ruta_json = os.path.join(carpeta_logs, "REPORTE_DETALLADO.json")
    exportar_json_detallado(resultados, ruta_json, metadatos={
        "plantilla": ruta_plantilla,
        "carpeta_trabajos": ruta_trabajos,
        "modo": modo.value,
        "seccion": seccion
    })
    logger.info(f"Reporte JSON generado en: {ruta_json}")

    # Excel Log Formateado
    ruta_xlsx_fija = os.path.join(carpeta_logs, "LOG_NOTAS.xlsx")
    ruta_xlsx_fecha = os.path.join(carpeta_logs, f"{seccion}_LOG_NOTAS_{fecha_str}.xlsx")
    exportar_excel_log(resultados, ruta_xlsx_fija, seccion=seccion)
    exportar_excel_log(resultados, ruta_xlsx_fecha, seccion=seccion)
    logger.info(f"Reporte Excel generado en: {ruta_xlsx_fija}")

    # Reporte de Observaciones en Texto (para copiar y pegar fácilmente)
    ruta_obs_fija = os.path.join(carpeta_logs, "OBSERVACIONES_ESTUDIANTES.txt")
    ruta_obs_fecha = os.path.join(carpeta_logs, f"{seccion}_OBSERVACIONES_{fecha_str}.txt")
    exportar_observaciones_txt(resultados, ruta_obs_fija, seccion=seccion)
    exportar_observaciones_txt(resultados, ruta_obs_fecha, seccion=seccion)
    logger.info(f"Reporte de Observaciones en texto generado en: {ruta_obs_fija}")

    if callback_progreso:
        callback_progreso({
            "evento": "completado",
            "total_archivos": total_archivos,
            "reporte_csv": ruta_csv,
            "reporte_json": ruta_json,
            "reporte_xlsx": ruta_xlsx_fija,
            "reporte_obs_txt": ruta_obs_fija,
        })

    logger.info("=== AUDITORÍA COMPLETADA CON ÉXITO ===")
    return resultados
