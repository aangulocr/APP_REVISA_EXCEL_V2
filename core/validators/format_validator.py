# -*- coding: utf-8 -*-
# ============================================================================
# core/validators/format_validator.py — Validador agnóstico de formatos numéricos
# ============================================================================
"""
Validador para 'formato_numero'.
Comprueba si las celdas tienen formato de Moneda, Porcentaje, Fecha o
Número Decimal sin depender de particularidades regionales de Excel.
"""

import re
from typing import List, Optional
import openpyxl.styles.numbers as ox_numbers
from core.validators.base import BaseValidator
from core.models import CriterioRubrica, ResultadoCriterio


def clasificar_formato(number_format: Optional[str]) -> str:
    """
    Clasifica un código de formato de Excel en una categoría general:
      'moneda', 'porcentaje', 'fecha', 'decimal', 'entero', 'texto' o 'general'.
    """
    if not number_format:
        return "general"

    nf = str(number_format).strip()
    nf_upper = nf.upper()

    if nf_upper in ("GENERAL", "@"):
        return "general"

    # 1. Porcentaje
    if "%" in nf:
        return "porcentaje"

    # 2. Moneda
    if any(simbolo in nf for simbolo in ("$", "€", "₡", "£", "¥", "Q", "L", "C$", "RD$")) or "[$" in nf:
        return "moneda"

    # 3. Fecha
    try:
        if ox_numbers.is_date_format(nf):
            return "fecha"
    except Exception:
        pass
    if any(k in nf_upper for k in ("YYYY", "AAAA", "YY", "DD", "MM/DD", "DD/MM")):
        return "fecha"

    # 4. Decimales
    if re.search(r"[0#][\.,][0#]+", nf):
        return "decimal"

    # 5. Entero
    if "#,##0" in nf or "0" in nf:
        return "entero"

    return "otro"


class FormatValidator(BaseValidator):
    """Valida la categoría semántica del formato numérico."""

    def validar(
        self,
        ws_estudiante,
        criterio: CriterioRubrica,
        ws_estudiante_data=None,
        ws_plantilla=None,
        ws_plantilla_data=None
    ) -> ResultadoCriterio:
        params = criterio.parametros or {}
        coords = self.obtener_coordenadas_rango(criterio.rango)

        if not coords:
            return ResultadoCriterio(
                criterio_id=criterio.id,
                descripcion=criterio.criterio,
                aprobado=False,
                puntos_obtenidos=0.0,
                puntos_maximos=criterio.puntos,
                mensaje_detalle=f"Rango '{criterio.rango}' no especificado.",
                obligatorio=criterio.obligatorio,
                tipo_validacion=criterio.tipo_validacion
            )

        formato_esperado = params.get("formato", params.get("tipo", "moneda")).lower().strip()
        if formato_esperado in ("currency", "dinero", "monetario"):
            formato_esperado = "moneda"
        elif formato_esperado in ("percentage", "pct", "%"):
            formato_esperado = "porcentaje"
        elif formato_esperado in ("date", "fechas"):
            formato_esperado = "fecha"
        elif formato_esperado in ("number", "numerico", "numérico", "coma"):
            formato_esperado = "decimal"
        elif formato_esperado in ("int", "integer"):
            formato_esperado = "entero"

        ignorar_vacias = params.get("ignorar_vacias", True)
        fallos: List[str] = []
        celdas_erroneas: List[str] = []

        for coord in coords:
            celda = ws_estudiante[coord] if ws_estudiante else None
            if celda is None:
                continue

            # Si la celda no tiene valor ni fórmula, omitir para no penalizar celdas vacías en un rango
            if ignorar_vacias and (celda.value is None or str(celda.value).strip() == ""):
                continue

            formato_celda = celda.number_format
            cat = clasificar_formato(formato_celda)

            if cat != formato_esperado:
                celdas_erroneas.append(coord)
                fallos.append(f"Celda {coord}: formato '{formato_celda}' clasificado como '{cat}' (esperaba '{formato_esperado}')")

        aprobado = len(fallos) == 0
        puntos = criterio.puntos if aprobado else 0.0
        mensaje = f"Formato '{formato_esperado}' correcto." if aprobado else "; ".join(fallos[:3])
        if len(fallos) > 3:
            mensaje += f" (y {len(fallos)-3} celdas más con formato incorrecto)"

        return ResultadoCriterio(
            criterio_id=criterio.id,
            descripcion=criterio.criterio,
            aprobado=aprobado,
            puntos_obtenidos=puntos,
            puntos_maximos=criterio.puntos,
            mensaje_detalle=mensaje,
            celdas_afectadas=celdas_erroneas if not aprobado else coords[:1],
            obligatorio=criterio.obligatorio,
            tipo_validacion=criterio.tipo_validacion
        )
