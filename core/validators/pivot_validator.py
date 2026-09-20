# -*- coding: utf-8 -*-
# ============================================================================
# core/validators/pivot_validator.py — Validador de Tablas Dinámicas
# ============================================================================
"""
Validador para 'tabla_dinamica'.
Verifica la existencia y configuración de Tablas Dinámicas (Pivot Tables),
tanto mediante inspección de estructuras openpyxl como vía COM en Windows.
"""

import logging
from typing import List, Dict, Any, Optional
from core.validators.base import BaseValidator
from core.models import CriterioRubrica, ResultadoCriterio

logger = logging.getLogger(__name__)


class PivotValidator(BaseValidator):
    """Valida la presencia y campos clave de Tablas Dinámicas."""

    def validar(
        self,
        ws_estudiante,
        criterio: CriterioRubrica,
        ws_estudiante_data=None,
        ws_plantilla=None,
        ws_plantilla_data=None
    ) -> ResultadoCriterio:
        params = criterio.parametros or {}
        campos_fila_req = [str(f).strip().upper() for f in params.get("campos_fila", [])]
        campos_col_req = [str(f).strip().upper() for f in params.get("campos_columna", [])]
        campos_valor_req = [str(f).strip().upper() for f in params.get("campos_valor", [])]
        campos_filtro_req = [str(f).strip().upper() for f in params.get("campos_filtro", [])]

        primera_celda = self.obtener_primera_celda(criterio.rango or "A3")

        if not ws_estudiante:
            return ResultadoCriterio(
                criterio_id=criterio.id,
                descripcion=criterio.criterio,
                aprobado=False,
                puntos_obtenidos=0.0,
                puntos_maximos=criterio.puntos,
                mensaje_detalle="Hoja no encontrada.",
                celdas_afectadas=[primera_celda],
                obligatorio=criterio.obligatorio,
                tipo_validacion=criterio.tipo_validacion
            )

        # 1. Detección en openpyxl: _pivots o relaciones XML
        tiene_pivot_openpyxl = False
        pivots_openpyxl = getattr(ws_estudiante, "_pivots", [])
        if pivots_openpyxl:
            tiene_pivot_openpyxl = True
        else:
            # Buscar en relaciones de la hoja (pivotTable*.xml)
            try:
                rels = getattr(ws_estudiante, "_rels", [])
                for rel in rels:
                    target = getattr(rel, "target", "")
                    rel_type = getattr(rel, "type", "")
                    if "pivotTable" in target or "pivotTable" in rel_type:
                        tiene_pivot_openpyxl = True
                        break
            except Exception:
                pass

        # 2. Si no hay COM y openpyxl no detectó pivot explícito, verificar si hay datos agregados en la hoja
        # o si la hoja se llama TablaDinamica o similar
        nombre_hoja_upper = ws_estudiante.title.upper()

        if not tiene_pivot_openpyxl:
            return ResultadoCriterio(
                criterio_id=criterio.id,
                descripcion=criterio.criterio,
                aprobado=False,
                puntos_obtenidos=0.0,
                puntos_maximos=criterio.puntos,
                mensaje_detalle=f"No se detectó Tabla Dinámica en la hoja '{ws_estudiante.title}'.",
                celdas_afectadas=[primera_celda],
                obligatorio=criterio.obligatorio,
                tipo_validacion=criterio.tipo_validacion
            )

        # Si los parámetros especifican campos particulares y se cuenta con información inyectada desde COM
        com_info = params.get("_com_pivot_info")
        detalles_campos = []

        if com_info and isinstance(com_info, dict):
            filas_est = [f.upper() for f in com_info.get("campos_fila", [])]
            cols_est = [f.upper() for f in com_info.get("campos_columna", [])]
            vals_est = [f.upper() for f in com_info.get("campos_valor", [])]
            filts_est = [f.upper() for f in com_info.get("campos_filtro", [])]

            for cf in campos_fila_req:
                if not any(cf in f for f in filas_est):
                    detalles_campos.append(f"Falta campo de fila '{cf}'")
            for cc in campos_col_req:
                if not any(cc in c for c in cols_est):
                    detalles_campos.append(f"Falta campo de columna '{cc}'")
            for cv in campos_valor_req:
                if not any(cv in v for v in vals_est):
                    detalles_campos.append(f"Falta campo de valor '{cv}'")
            for cfilt in campos_filtro_req:
                if not any(cfilt in fi for fi in filts_est):
                    detalles_campos.append(f"Falta campo de filtro '{cfilt}'")

        if detalles_campos:
            return ResultadoCriterio(
                criterio_id=criterio.id,
                descripcion=criterio.criterio,
                aprobado=False,
                puntos_obtenidos=0.0,
                puntos_maximos=criterio.puntos,
                mensaje_detalle="Tabla Dinámica encontrada pero con campos incompletos: " + "; ".join(detalles_campos),
                celdas_afectadas=[primera_celda],
                obligatorio=criterio.obligatorio,
                tipo_validacion=criterio.tipo_validacion
            )

        return ResultadoCriterio(
            criterio_id=criterio.id,
            descripcion=criterio.criterio,
            aprobado=True,
            puntos_obtenidos=criterio.puntos,
            puntos_maximos=criterio.puntos,
            mensaje_detalle="Tabla Dinámica verificada correctamente.",
            celdas_afectadas=[primera_celda],
            obligatorio=criterio.obligatorio,
            tipo_validacion=criterio.tipo_validacion
        )
