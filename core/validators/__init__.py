# -*- coding: utf-8 -*-
# ============================================================================
# core/validators/__init__.py — Catálogo y fábrica de validadores
# ============================================================================

from typing import Dict, Type
from core.validators.base import BaseValidator
from core.validators.value_validator import ValueValidator
from core.validators.formula_validator import FormulaValidator
from core.validators.format_validator import FormatValidator
from core.validators.style_validator import StyleValidator
from core.validators.structure_validator import StructureValidator
from core.validators.pivot_validator import PivotValidator

# Mapeo de identificadores de validación a sus clases
_REGISTRO_VALIDADORES: Dict[str, Type[BaseValidator]] = {
    "valor_exacto": ValueValidator,
    "valor_numerico": ValueValidator,
    "formula_flexible": FormulaValidator,
    "formato_numero": FormatValidator,
    "estilo_visual_flexible": StyleValidator,
    "celdas_combinadas": StructureValidator,
    "validacion_datos": StructureValidator,
    "tabla_dinamica": PivotValidator,
}


def obtener_validador(tipo_validacion: str) -> BaseValidator:
    """
    Instancia y retorna el validador correspondiente según el tipo solicitado.
    Si no se reconoce, utiliza ValueValidator por defecto.
    """
    clave = (tipo_validacion or "").lower().strip()
    clase_validador = _REGISTRO_VALIDADORES.get(clave, ValueValidator)
    return clase_validador()


__all__ = [
    "BaseValidator",
    "ValueValidator",
    "FormulaValidator",
    "FormatValidator",
    "StyleValidator",
    "StructureValidator",
    "PivotValidator",
    "obtener_validador",
]
