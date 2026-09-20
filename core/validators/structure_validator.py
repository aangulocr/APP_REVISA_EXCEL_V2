# -*- coding: utf-8 -*-
# ============================================================================
# core/validators/structure_validator.py — Validador de celdas combinadas y validaciones
# ============================================================================
"""
Validador para 'celdas_combinadas' y 'validacion_datos'.
Verifica que los rangos especificados hayan sido combinados (merged) o cuenten
con reglas de validación de datos (listas desplegables, enteros, etc.).
"""

from typing import List, Any
import openpyxl.utils
from openpyxl.worksheet.cell_range import CellRange
from core.validators.base import BaseValidator
from core.models import CriterioRubrica, ResultadoCriterio


class StructureValidator(BaseValidator):
    """Valida celdas combinadas y validación de datos."""

    def validar(
        self,
        ws_estudiante,
        criterio: CriterioRubrica,
        ws_estudiante_data=None,
        ws_plantilla=None,
        ws_plantilla_data=None
    ) -> ResultadoCriterio:
        tipo = criterio.tipo_validacion.lower().strip()
        if tipo == "celdas_combinadas":
            return self._validar_combinadas(ws_estudiante, criterio)
        elif tipo == "validacion_datos":
            return self._validar_validacion_datos(ws_estudiante, criterio)
        else:
            return ResultadoCriterio(
                criterio_id=criterio.id,
                descripcion=criterio.criterio,
                aprobado=False,
                puntos_obtenidos=0.0,
                puntos_maximos=criterio.puntos,
                mensaje_detalle=f"Tipo de validación estructural no soportado: '{tipo}'",
                obligatorio=criterio.obligatorio,
                tipo_validacion=criterio.tipo_validacion
            )

    def _validar_combinadas(self, ws_estudiante, criterio: CriterioRubrica) -> ResultadoCriterio:
        rango_req = criterio.rango.strip().upper()
        if not ws_estudiante:
            return ResultadoCriterio(
                criterio_id=criterio.id,
                descripcion=criterio.criterio,
                aprobado=False,
                puntos_obtenidos=0.0,
                puntos_maximos=criterio.puntos,
                mensaje_detalle="Hoja no disponible.",
                obligatorio=criterio.obligatorio,
                tipo_validacion=criterio.tipo_validacion
            )

        rangos_combinados = ws_estudiante.merged_cells.ranges
        coincide = False
        rango_encontrado = None

        try:
            rango_req_obj = CellRange(rango_req)
            for m in rangos_combinados:
                # Coincidencia exacta de rango o contención
                if m.coord == rango_req:
                    coincide = True
                    rango_encontrado = m.coord
                    break
                if rango_req_obj.issubset(m):
                    coincide = True
                    rango_encontrado = m.coord
                    break
        except Exception:
            # Fallback a comparación de strings
            for m in rangos_combinados:
                if str(m).upper() == rango_req:
                    coincide = True
                    rango_encontrado = str(m)
                    break

        primera = self.obtener_primera_celda(rango_req)
        if coincide:
            return ResultadoCriterio(
                criterio_id=criterio.id,
                descripcion=criterio.criterio,
                aprobado=True,
                puntos_obtenidos=criterio.puntos,
                puntos_maximos=criterio.puntos,
                mensaje_detalle=f"Rango combinado correctamente ({rango_encontrado}).",
                celdas_afectadas=[primera],
                obligatorio=criterio.obligatorio,
                tipo_validacion=criterio.tipo_validacion
            )
        else:
            return ResultadoCriterio(
                criterio_id=criterio.id,
                descripcion=criterio.criterio,
                aprobado=False,
                puntos_obtenidos=0.0,
                puntos_maximos=criterio.puntos,
                mensaje_detalle=f"El rango '{rango_req}' no está combinado (merged).",
                celdas_afectadas=[primera],
                obligatorio=criterio.obligatorio,
                tipo_validacion=criterio.tipo_validacion
            )

    def _validar_validacion_datos(self, ws_estudiante, criterio: CriterioRubrica) -> ResultadoCriterio:
        rango_req = criterio.rango.strip().upper()
        coords_req = self.obtener_coordenadas_rango(rango_req)
        primera = self.obtener_primera_celda(rango_req)

        if not ws_estudiante:
            return ResultadoCriterio(
                criterio_id=criterio.id,
                descripcion=criterio.criterio,
                aprobado=False,
                puntos_obtenidos=0.0,
                puntos_maximos=criterio.puntos,
                mensaje_detalle="Hoja no disponible.",
                obligatorio=criterio.obligatorio,
                tipo_validacion=criterio.tipo_validacion
            )

        dvs = ws_estudiante.data_validations.dataValidation if ws_estudiante.data_validations else []
        if not dvs:
            return ResultadoCriterio(
                criterio_id=criterio.id,
                descripcion=criterio.criterio,
                aprobado=False,
                puntos_obtenidos=0.0,
                puntos_maximos=criterio.puntos,
                mensaje_detalle=f"No se encontró ninguna validación de datos en la hoja (se esperaba en '{rango_req}').",
                celdas_afectadas=[primera],
                obligatorio=criterio.obligatorio,
                tipo_validacion=criterio.tipo_validacion
            )

        params = criterio.parametros or {}
        tipo_esperado = params.get("tipo", "").lower().strip()

        # Buscar si alguna validación de datos cubre las celdas requeridas
        encontrada = False
        detalles_dv = ""

        for dv in dvs:
            cubre_rango = False
            # Las validaciones contienen rangos en sqref
            sqref_str = str(dv.sqref).upper()
            if rango_req in sqref_str:
                cubre_rango = True
            else:
                # Comprobar si las celdas individuales están dentro
                for r in dv.sqref.ranges:
                    for coord in coords_req:
                        try:
                            c_col, c_row = openpyxl.utils.coordinate_to_tuple(coord)
                            if (r.min_row <= c_row <= r.max_row) and (r.min_col <= c_col <= r.max_col):
                                cubre_rango = True
                                break
                        except Exception:
                            pass
                    if cubre_rango:
                        break

            if cubre_rango:
                if tipo_esperado and dv.type and dv.type.lower() != tipo_esperado:
                    detalles_dv = f"Tipo de validación es '{dv.type}' (se esperaba '{tipo_esperado}')"
                else:
                    encontrada = True
                    detalles_dv = f"Tipo: {dv.type or 'lista/valor'}"
                    break

        if encontrada:
            return ResultadoCriterio(
                criterio_id=criterio.id,
                descripcion=criterio.criterio,
                aprobado=True,
                puntos_obtenidos=criterio.puntos,
                puntos_maximos=criterio.puntos,
                mensaje_detalle=f"Validación de datos encontrada en '{rango_req}' ({detalles_dv}).",
                celdas_afectadas=[primera],
                obligatorio=criterio.obligatorio,
                tipo_validacion=criterio.tipo_validacion
            )
        else:
            msg = f"Validación de datos ausente o incorrecta en '{rango_req}'."
            if detalles_dv:
                msg += f" {detalles_dv}"
            return ResultadoCriterio(
                criterio_id=criterio.id,
                descripcion=criterio.criterio,
                aprobado=False,
                puntos_obtenidos=0.0,
                puntos_maximos=criterio.puntos,
                mensaje_detalle=msg,
                celdas_afectadas=[primera],
                obligatorio=criterio.obligatorio,
                tipo_validacion=criterio.tipo_validacion
            )
