# -*- coding: utf-8 -*-
# ============================================================================
# core/validators/style_validator.py — Validador flexible de estilos visuales
# ============================================================================
"""
Validador para 'estilo_visual_flexible'.
Evalúa atributos de diseño visual (negrita, relleno, bordes, alineación)
sin rigidez cromática extrema, valorando la intención estética del estudiante.
"""

from typing import List, Optional
from core.validators.base import BaseValidator
from core.models import CriterioRubrica, ResultadoCriterio


class StyleValidator(BaseValidator):
    """Valida estilos visuales (negrita, relleno, bordes, alineación)."""

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
                mensaje_detalle=f"Rango '{criterio.rango}' inválido.",
                obligatorio=criterio.obligatorio,
                tipo_validacion=criterio.tipo_validacion
            )

        req_negrita = bool(params.get("requiere_negrita", params.get("negrita", False)))
        req_relleno = bool(params.get("requiere_relleno", params.get("relleno", False)))
        req_bordes = bool(params.get("requiere_bordes", params.get("bordes", False)))
        req_alineacion = params.get("alineacion", params.get("alineacion_h", "")).lower().strip()

        fallos: List[str] = []
        celdas_erroneas: List[str] = []

        for coord in coords:
            celda = ws_estudiante[coord] if ws_estudiante else None
            if celda is None:
                continue

            errores_celda = []

            # 1. Negrita
            if req_negrita:
                es_negrita = bool(celda.font and celda.font.bold)
                if not es_negrita:
                    errores_celda.append("falta negrita")

            # 2. Relleno
            if req_relleno:
                tiene_relleno = False
                if celda.fill and celda.fill.fill_type not in (None, "none"):
                    # Verificar que no sea blanco puro o transparente
                    fg = celda.fill.fgColor
                    if fg:
                        rgb = str(fg.rgb or "")
                        # Si no es blanco o vacío
                        if rgb not in ("00000000", "FFFFFFFF", "") or fg.theme is not None or fg.indexed is not None:
                            tiene_relleno = True
                    else:
                        tiene_relleno = True
                if not tiene_relleno:
                    errores_celda.append("falta color de relleno")

            # 3. Bordes
            if req_bordes:
                tiene_bordes = False
                b = celda.border
                if b and (b.left.style or b.right.style or b.top.style or b.bottom.style):
                    tiene_bordes = True
                if not tiene_bordes:
                    errores_celda.append("falta aplicar bordes")

            # 4. Alineación
            if req_alineacion:
                ali_actual = (celda.alignment.horizontal or "general").lower() if celda.alignment else "general"
                if req_alineacion == "centrado":
                    req_alineacion = "center"
                if ali_actual != req_alineacion:
                    errores_celda.append(f"alineación horizontal es '{ali_actual}' (esperaba '{req_alineacion}')")

            if errores_celda:
                celdas_erroneas.append(coord)
                fallos.append(f"Celda {coord}: {', '.join(errores_celda)}")

        aprobado = len(fallos) == 0
        puntos = criterio.puntos if aprobado else 0.0
        mensaje = "Estilo visual correcto." if aprobado else "; ".join(fallos[:3])
        if len(fallos) > 3:
            mensaje += f" (y {len(fallos)-3} celdas más con diferencias de estilo)"

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
