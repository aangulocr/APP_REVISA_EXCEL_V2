# -*- coding: utf-8 -*-
# ============================================================================
# core/evaluators/base_evaluator.py — Interfaz base y utilidades de marcado visual
# ============================================================================
"""
Define la interfaz común para los motores de evaluación (Modo A y Modo B),
así como las funciones para aplicar resaltado cromático suave y comentarios
de retroalimentación en los libros evaluados.
"""

from abc import ABC, abstractmethod
from typing import Optional, List
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment

from core.models import ResultadoEstudiante
from core.config import COLOR_ERROR_FILL, COLOR_SUCCESS_FILL


# Rellenos cromáticos
RELLENO_ERROR = PatternFill(
    start_color=COLOR_ERROR_FILL,
    end_color=COLOR_ERROR_FILL,
    fill_type="solid"
)

RELLENO_EXITO = PatternFill(
    start_color=COLOR_SUCCESS_FILL,
    end_color=COLOR_SUCCESS_FILL,
    fill_type="solid"
)


def aplicar_marcado_error(celda, mensaje: str, autor: str = "Auditor Excel V2"):
    """
    Marca una celda con relleno rojo suave y agrega comentario explicativo.
    """
    if celda is None:
        return
    celda.fill = RELLENO_ERROR

    comentario_previo = celda.comment.text if celda.comment else ""
    nuevo_texto = f"{comentario_previo}\n{mensaje}".strip() if comentario_previo else mensaje

    celda.comment = Comment(text=nuevo_texto, author=autor)
    celda.comment.width = 380
    celda.comment.height = 110 + (nuevo_texto.count("\n") * 24)


def aplicar_marcado_aprobado(celda, mensaje: str, autor: str = "Auditor Excel V2"):
    """
    Inserta un comentario positivo en verde/éxito en la celda del estudiante.
    """
    if celda is None:
        return
    comentario_previo = celda.comment.text if celda.comment else ""
    nuevo_texto = f"{comentario_previo}\n{mensaje}".strip() if comentario_previo else mensaje

    celda.comment = Comment(text=nuevo_texto, author=autor)
    celda.comment.width = 360
    celda.comment.height = 90


class BaseEvaluator(ABC):
    """Clase base para los evaluadores Modo A y Modo B."""

    @abstractmethod
    def evaluar(
        self,
        ruta_plantilla: str,
        ruta_estudiante: str,
        carpeta_salida_revisados: str
    ) -> ResultadoEstudiante:
        """
        Ejecuta la auditoría completa sobre un archivo de estudiante
        y genera el archivo revisado (_rev.xlsx).
        """
        pass
