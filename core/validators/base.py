# -*- coding: utf-8 -*-
# ============================================================================
# core/validators/base.py — Clase base abstracta para validadores de rúbrica
# ============================================================================
"""
Define la interfaz BaseValidator y métodos auxiliares para la resolución
de rangos de celdas, extracción de coordenadas y diagnóstico de errores.
"""

from abc import ABC, abstractmethod
from typing import List, Tuple, Optional, Any
import openpyxl.utils

from core.models import CriterioRubrica, ResultadoCriterio


class BaseValidator(ABC):
    """Clase base de la cual heredan todos los validadores de criterios."""

    @abstractmethod
    def validar(
        self,
        ws_estudiante,
        criterio: CriterioRubrica,
        ws_estudiante_data=None,
        ws_plantilla=None,
        ws_plantilla_data=None
    ) -> ResultadoCriterio:
        """
        Ejecuta la validación del criterio sobre la hoja del estudiante.

        Parámetros:
            ws_estudiante: Worksheet de openpyxl con fórmulas (data_only=False).
            criterio: Objeto CriterioRubrica a evaluar.
            ws_estudiante_data: Worksheet opcional con valores calculados (data_only=True).
            ws_plantilla: Worksheet opcional de la plantilla (fórmulas).
            ws_plantilla_data: Worksheet opcional de la plantilla (valores).

        Retorna:
            ResultadoCriterio con el diagnóstico y puntaje obtenido.
        """
        pass

    @staticmethod
    def obtener_coordenadas_rango(rango_str: str) -> List[str]:
        """
        Dada una cadena de rango (ej. 'D20' o 'D2:D19' o 'A1,B2'),
        retorna la lista de coordenadas individuales en orden.
        """
        if not rango_str:
            return []
        coords = []
        partes = [p.strip().replace("$", "") for p in rango_str.split(",") if p.strip()]
        for parte in partes:
            if ":" in parte:
                try:
                    min_col, min_row, max_col, max_row = openpyxl.utils.range_boundaries(parte)
                    if min_col and min_row and max_col and max_row:
                        for r in range(min_row, max_row + 1):
                            for c in range(min_col, max_col + 1):
                                let = openpyxl.utils.get_column_letter(c)
                                coords.append(f"{let}{r}")
                except Exception:
                    coords.append(parte)
            else:
                coords.append(parte)
        return coords

    @staticmethod
    def obtener_primera_celda(rango_str: str) -> str:
        """Retorna la primera celda de un rango (ej: 'D2:D19' -> 'D2')."""
        if not rango_str:
            return "A1"
        coords = BaseValidator.obtener_coordenadas_rango(rango_str)
        return coords[0] if coords else rango_str.split(":")[0].strip()
