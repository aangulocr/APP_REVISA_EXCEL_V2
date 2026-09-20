# -*- coding: utf-8 -*-
# ============================================================================
# core/models.py — Modelos de datos tipados para APP_REVISA_EXCEL_V2
# ============================================================================
"""
Define las estructuras de datos y tipos estáticos que representan
rúbricas, criterios, resultados de validación y resúmenes de auditoría.
"""

from enum import Enum
from typing import Dict, List, Optional, Any
from dataclasses import dataclass, field
from datetime import datetime


class ModoEvaluacion(str, Enum):
    MODO_A_ESPEJO = "modo_a_espejo"
    MODO_B_RUBRICA = "modo_b_rubrica"

    @classmethod
    def from_str(cls, val: str) -> "ModoEvaluacion":
        val_lower = val.lower().strip()
        if "rubrica" in val_lower or "b" in val_lower:
            return cls.MODO_B_RUBRICA
        return cls.MODO_A_ESPEJO


@dataclass
class CriterioRubrica:
    """Representa una regla declarativa de evaluación."""
    id: str
    criterio: str
    hoja: str
    rango: str
    tipo_validacion: str
    parametros: Dict[str, Any] = field(default_factory=dict)
    puntos: float = 10.0
    obligatorio: bool = False

    def to_dict(self) -> Dict[str, Any]:
        return {
            "id": self.id,
            "criterio": self.criterio,
            "hoja": self.hoja,
            "rango": self.rango,
            "tipo_validacion": self.tipo_validacion,
            "parametros": self.parametros,
            "puntos": self.puntos,
            "obligatorio": self.obligatorio,
        }


@dataclass
class ResultadoCriterio:
    """Resultado de la evaluación de un criterio específico para un estudiante."""
    criterio_id: str
    descripcion: str
    aprobado: bool
    puntos_obtenidos: float
    puntos_maximos: float
    mensaje_detalle: str = ""
    celdas_afectadas: List[str] = field(default_factory=list)
    obligatorio: bool = False
    tipo_validacion: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return {
            "criterio_id": self.criterio_id,
            "descripcion": self.descripcion,
            "aprobado": self.aprobado,
            "puntos_obtenidos": self.puntos_obtenidos,
            "puntos_maximos": self.puntos_maximos,
            "mensaje_detalle": self.mensaje_detalle,
            "celdas_afectadas": self.celdas_afectadas,
            "obligatorio": self.obligatorio,
            "tipo_validacion": self.tipo_validacion,
        }


@dataclass
class ResultadoHoja:
    """Resumen de evaluación para una hoja de cálculo individual (Modo A o Modo B)."""
    hoja: str
    hoja_estudiante: Optional[str] = None
    aciertos: int = 0
    errores: int = 0
    porcentaje: float = 0.0
    detalles: List[str] = field(default_factory=list)
    desplazamiento: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        return {
            "hoja": self.hoja,
            "hoja_estudiante": self.hoja_estudiante,
            "aciertos": self.aciertos,
            "errores": self.errores,
            "porcentaje": self.porcentaje,
            "detalles": self.detalles,
            "desplazamiento": self.desplazamiento,
        }


@dataclass
class ResultadoEstudiante:
    """Resultado integral de la auditoría de un libro de estudiante."""
    archivo: str
    nombre_estudiante: str
    modo: str
    puntos_obtenidos: float = 0.0
    puntos_totales: float = 100.0
    nota_final_100: float = 0.0
    aprobado: bool = False
    criterios_evaluados: List[ResultadoCriterio] = field(default_factory=list)
    criterios_fallados: List[str] = field(default_factory=list)
    detalle_hojas: List[ResultadoHoja] = field(default_factory=list)
    advertencias: List[str] = field(default_factory=list)
    errores_com: List[str] = field(default_factory=list)
    total_aciertos: int = 0
    total_errores: int = 0
    archivo_revisado: Optional[str] = None
    tiempo_segundos: float = 0.0
    fecha_hora: str = field(default_factory=lambda: datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    def to_dict(self) -> Dict[str, Any]:
        return {
            "archivo": self.archivo,
            "nombre_estudiante": self.nombre_estudiante,
            "modo": self.modo,
            "puntos_obtenidos": round(self.puntos_obtenidos, 2),
            "puntos_totales": round(self.puntos_totales, 2),
            "nota_final_100": round(self.nota_final_100, 2),
            "aprobado": self.aprobado,
            "criterios_evaluados": [c.to_dict() for c in self.criterios_evaluados],
            "criterios_fallados": self.criterios_fallados,
            "detalle_hojas": [h.to_dict() for h in self.detalle_hojas],
            "advertencias": self.advertencias,
            "errores_com": self.errores_com,
            "total_aciertos": self.total_aciertos,
            "total_errores": self.total_errores,
            "archivo_revisado": self.archivo_revisado,
            "tiempo_segundos": round(self.tiempo_segundos, 2),
            "fecha_hora": self.fecha_hora,
        }
