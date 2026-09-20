# -*- coding: utf-8 -*-
# ============================================================================
# core/shift_detector.py — Detección heurística de desplazamiento de celdas
# ============================================================================
"""
Analiza si un archivo de estudiante tiene un desplazamiento sistemático
(offset) de filas o columnas respecto a la plantilla (por ejemplo, haber
insertado una fila o columna en blanco al principio).
"""

from typing import Tuple, Optional, Dict, Any
from openpyxl.worksheet.worksheet import Worksheet


def detectar_desplazamiento(
    ws_plantilla: Worksheet,
    ws_estudiante: Worksheet,
    max_muestra_filas: int = 25,
    max_muestra_cols: int = 15
) -> Tuple[Optional[Tuple[int, int]], float, str]:
    """
    Evalúa si las celdas del estudiante parecen estar desplazadas por (delta_row, delta_col).

    Retorna:
        ((delta_row, delta_col), confianza, mensaje_advertencia)
        Si no hay desplazamiento o la coincidencia directa es buena, retorna (None, 0.0, "")
    """
    if ws_plantilla is None or ws_estudiante is None:
        return None, 0.0, ""

    # Hojas de gráfico independientes (Chartsheet) no poseen celdas ni filas
    if not hasattr(ws_plantilla, "max_row") or not hasattr(ws_estudiante, "max_row"):
        return None, 0.0, ""
    if not hasattr(ws_plantilla, "cell") or not hasattr(ws_estudiante, "cell"):
        return None, 0.0, ""

    # Recolectar celdas de muestra no vacías de la plantilla
    muestras = []
    max_r = min(ws_plantilla.max_row or 1, max_muestra_filas)
    max_c = min(ws_plantilla.max_column or 1, max_muestra_cols)

    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            val = ws_plantilla.cell(row=r, column=c).value
            if val is not None and str(val).strip() != "":
                muestras.append((r, c, str(val).strip().upper()))

    if len(muestras) < 5:
        return None, 0.0, ""

    # Calcular coincidencia directa (0, 0)
    coincidencias_directas = 0
    for r, c, val_p in muestras:
        val_e = ws_estudiante.cell(row=r, column=c).value
        val_e_str = str(val_e).strip().upper() if val_e is not None else ""
        if val_p == val_e_str:
            coincidencias_directas += 1

    ratio_directo = coincidencias_directas / len(muestras)

    # Si la coincidencia directa ya es alta (> 75%), no hay sospecha de desplazamiento
    if ratio_directo >= 0.75:
        return None, 0.0, ""

    # Probar offsets plausibles: filas -2 a +2, cols -1 a +1
    candidatos_offsets = [
        (1, 0),   # Fila adicional insertada arriba (muy común)
        (2, 0),   # Dos filas adicionales
        (-1, 0),  # Fila eliminada arriba
        (0, 1),   # Columna adicional a la izquierda
        (1, 1),   # Fila y columna adicionales
        (-1, 1),
        (0, -1),
    ]

    mejor_offset = None
    mejor_ratio = ratio_directo

    for dr, dc in candidatos_offsets:
        matches = 0
        total_evaluados = 0

        for r, c, val_p in muestras:
            nr = r + dr
            nc = c + dc
            if nr < 1 or nc < 1:
                continue
            total_evaluados += 1
            val_e = ws_estudiante.cell(row=nr, column=nc).value
            val_e_str = str(val_e).strip().upper() if val_e is not None else ""
            if val_p == val_e_str:
                matches += 1

        if total_evaluados > 0:
            ratio = matches / total_evaluados
            if ratio > mejor_ratio:
                mejor_ratio = ratio
                mejor_offset = (dr, dc)

    # Solo reportar si el offset produce una mejora sustancial (> 70% acierto y superando por al menos 30% al directo)
    if mejor_offset and mejor_ratio >= 0.70 and (mejor_ratio - ratio_directo) >= 0.30:
        dr, dc = mejor_offset
        partes = []
        if dr > 0:
            partes.append(f"+{dr} fila(s) abajo")
        elif dr < 0:
            partes.append(f"{dr} fila(s) arriba")
        if dc > 0:
            partes.append(f"+{dc} columna(s) a la derecha")
        elif dc < 0:
            partes.append(f"{dc} columna(s) a la izquierda")

        desc_offset = " y ".join(partes)
        msg = (
            f"⚠️ Posible desplazamiento detectado: los datos coinciden en un {mejor_ratio*100:.1f}% "
            f"con un desplazamiento de {desc_offset} (coincidencia directa era solo {ratio_directo*100:.1f}%). "
            f"Es probable que el estudiante haya insertado o borrado filas/columnas."
        )
        return mejor_offset, mejor_ratio, msg

    return None, 0.0, ""
