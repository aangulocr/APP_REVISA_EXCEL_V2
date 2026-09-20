# -*- coding: utf-8 -*-
# ============================================================================
# core/validators/chart_validator.py — Validador de Gráficos y Hojas de Gráfico
# ============================================================================
"""
Validador para 'grafico_dinamico', 'hoja_grafico' y 'grafico'.
Verifica que el gráfico o gráfico dinámico exista y haya sido movido como
Hoja de Gráfico (ChartSheet) o incrustado según los parámetros de la rúbrica.
"""

import logging
from typing import List, Dict, Any, Optional
from core.validators.base import BaseValidator
from core.models import CriterioRubrica, ResultadoCriterio

logger = logging.getLogger(__name__)


class ChartValidator(BaseValidator):
    """Valida la presencia y configuración de Gráficos Dinámicos y Hojas de Gráfico."""

    def validar(
        self,
        ws_estudiante,
        criterio: CriterioRubrica,
        ws_estudiante_data=None,
        ws_plantilla=None,
        ws_plantilla_data=None
    ) -> ResultadoCriterio:
        params = criterio.parametros or {}
        tipo_hoja_req = str(params.get("tipo_hoja", "")).lower().strip()
        origen_req = str(params.get("origen", "")).strip()

        if ws_estudiante is None:
            return ResultadoCriterio(
                criterio_id=criterio.id,
                descripcion=criterio.criterio,
                aprobado=False,
                puntos_obtenidos=0.0,
                puntos_maximos=criterio.puntos,
                mensaje_detalle=f"Hoja de gráfico '{criterio.hoja}' no encontrada en el libro.",
                celdas_afectadas=[],
                obligatorio=criterio.obligatorio,
                tipo_validacion=criterio.tipo_validacion
            )

        # 1. Detectar si es una Hoja de Gráfico (ChartSheet)
        # En openpyxl, una hoja movida a "Hoja nueva" es de clase Chartsheet
        clase_nombre = ws_estudiante.__class__.__name__
        tagname = getattr(ws_estudiante, "tagname", "")
        es_chartsheet = (clase_nombre == "Chartsheet" or tagname == "chartsheet")

        # 2. Detectar si tiene gráficos incrustados (si es hoja normal)
        charts = getattr(ws_estudiante, "_charts", [])
        tiene_graficos = bool(charts) or es_chartsheet

        # Si se requiere específicamente hoja de gráfico (Mover Gráfico -> Hoja nueva)
        # O si el criterio menciona "mover" / "hoja nueva" / tipo_hoja == "chartsheet"
        requiere_hoja_grafico = (
            tipo_hoja_req in ("chartsheet", "hoja_grafico", "hoja_nueva", "nueva_hoja")
            or "mover" in criterio.criterio.lower()
            or "hoja nueva" in criterio.criterio.lower()
        )

        if requiere_hoja_grafico and not es_chartsheet:
            return ResultadoCriterio(
                criterio_id=criterio.id,
                descripcion=criterio.criterio,
                aprobado=False,
                puntos_obtenidos=0.0,
                puntos_maximos=criterio.puntos,
                mensaje_detalle=(
                    f"La hoja '{criterio.hoja}' existe pero es una hoja de cálculo estándar. "
                    "El gráfico dinámico debía moverse como 'Hoja nueva' (ChartSheet) usando Mover Gráfico."
                ),
                celdas_afectadas=[],
                obligatorio=criterio.obligatorio,
                tipo_validacion=criterio.tipo_validacion
            )

        if not tiene_graficos and not es_chartsheet:
            return ResultadoCriterio(
                criterio_id=criterio.id,
                descripcion=criterio.criterio,
                aprobado=False,
                puntos_obtenidos=0.0,
                puntos_maximos=criterio.puntos,
                mensaje_detalle=f"No se detectó ningún gráfico dinámico en la hoja '{criterio.hoja}'.",
                celdas_afectadas=[],
                obligatorio=criterio.obligatorio,
                tipo_validacion=criterio.tipo_validacion
            )

        detalle = f"Gráfico dinámico detectado exitosamente en la hoja '{criterio.hoja}' (Hoja de Gráfico / ChartSheet)."
        if origen_req:
            detalle += f" Tabla Dinámica origen esperada: '{origen_req}'."

        return ResultadoCriterio(
            criterio_id=criterio.id,
            descripcion=criterio.criterio,
            aprobado=True,
            puntos_obtenidos=criterio.puntos,
            puntos_maximos=criterio.puntos,
            mensaje_detalle=detalle,
            celdas_afectadas=[],
            obligatorio=criterio.obligatorio,
            tipo_validacion=criterio.tipo_validacion
        )
