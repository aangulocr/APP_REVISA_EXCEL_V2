# -*- coding: utf-8 -*-
# ============================================================================
# core/rubric_parser.py — Parser inteligente de rúbricas (Excel / JSON)
# ============================================================================
"""
Carga e interpreta las reglas declarativas de evaluación desde:
  1. Hoja oculta o visible '_RUBRICA' dentro del archivo de plantilla .xlsx
  2. Archivo '<nombre_plantilla>.rubrica.json' o 'rubrica.json'
"""

import os
import re
import json
import logging
import unicodedata
from typing import List, Dict, Tuple, Optional, Any
import openpyxl

from core.models import CriterioRubrica
from core.config import (
    HOJA_RUBRICA_DEFAULT,
    NOMBRES_HOJA_RUBRICA,
    SUFIJO_ARCHIVO_RUBRICA,
    ARCHIVO_RUBRICA_GLOBAL,
    VALIDADORES_DISPONIBLES,
)
from core.sheet_matcher import es_rubrica, normalizar_texto

logger = logging.getLogger(__name__)


def _limpiar_json_str(texto: str) -> Dict[str, Any]:
    """
    Parsea de forma permisiva una cadena de parámetros en formato JSON.
    Acepta comillas simples, formatos no estrictos o cadenas vacías.
    """
    if not texto:
        return {}
    if isinstance(texto, dict):
        return texto

    texto_str = str(texto).strip()
    if not texto_str or texto_str.lower() in ("none", "null", "-", ""):
        return {}

    # Intentar parseo directo
    try:
        return json.loads(texto_str)
    except Exception:
        pass

    # Reemplazar comillas simples por dobles si aplica
    try:
        corregido = texto_str.replace("'", '"')
        # Si es true/false en minúsculas/mayúsculas
        corregido = re.sub(r'\bTrue\b', 'true', corregido)
        corregido = re.sub(r'\bFalse\b', 'false', corregido)
        return json.loads(corregido)
    except Exception as e:
        logger.warning(f"No se pudo parsear JSON en parámetros: '{texto_str}' -> {e}")
        return {"raw": texto_str}


def _normalizar_nombre_col(col_name: Any) -> str:
    """Normaliza encabezados de columna de la rúbrica."""
    if not col_name:
        return ""
    norm = normalizar_texto(str(col_name))
    # Quitar guiones y espacios
    norm = norm.replace("_", "").replace("-", "").replace(" ", "")
    return norm


def parsear_rubrica_desde_hoja(ws) -> Tuple[List[CriterioRubrica], float, List[str]]:
    """
    Lee una hoja de cálculo openpyxl y extrae los criterios de la rúbrica.

    Retorna:
        (criterios: List[CriterioRubrica], puntos_totales: float, advertencias: List[str])
    """
    criterios: List[CriterioRubrica] = []
    advertencias: List[str] = []

    if ws is None:
        return criterios, 0.0, ["La hoja de rúbrica es nula."]

    # Buscar fila de encabezados (revisar filas 1 a 5)
    header_row = None
    col_map: Dict[str, int] = {}

    mapeo_esperado = {
        "id": ["id", "codigo", "num", "criterioid"],
        "criterio": ["criterio", "descripcion", "detalle", "nombre"],
        "hoja": ["hoja", "sheet", "pestana"],
        "rango": ["rango", "celda", "celdas", "range"],
        "tipo_validacion": ["tipovalidacion", "tipo", "validador", "validacion"],
        "parametros": ["parametros", "params", "argumentos", "config"],
        "puntos": ["puntos", "puntaje", "peso", "valor", "pts"],
        "obligatorio": ["obligatorio", "esobligatorio", "requerido"],
    }

    max_r = min(ws.max_row or 1, 10)
    max_c = min(ws.max_column or 1, 20)

    for r in range(1, max_r + 1):
        fila_cols = {}
        for c in range(1, max_c + 1):
            val = ws.cell(row=r, column=c).value
            if val:
                norm = _normalizar_nombre_col(val)
                for clave_interna, sinonimos in mapeo_esperado.items():
                    if norm in sinonimos:
                        fila_cols[clave_interna] = c
                        break

        # Si encontramos al menos criterio, hoja, rango o id
        if "criterio" in fila_cols and ("rango" in fila_cols or "hoja" in fila_cols):
            header_row = r
            col_map = fila_cols
            break

    if header_row is None:
        return criterios, 0.0, [
            f"No se detectaron encabezados válidos en la hoja de rúbrica '{ws.title}'. "
            "Se esperan columnas como: ID, CRITERIO, HOJA, RANGO, TIPO_VALIDACION, PARAMETROS, PUNTOS."
        ]

    # Leer filas de datos
    fila_actual = header_row + 1
    total_filas = ws.max_row or fila_actual

    for r in range(fila_actual, total_filas + 1):
        criterio_desc = ws.cell(row=r, column=col_map.get("criterio", 2)).value
        if criterio_desc is None or str(criterio_desc).strip() == "":
            continue

        c_id = ws.cell(row=r, column=col_map.get("id", 1)).value
        if not c_id:
            c_id = f"CRIT_{len(criterios) + 1:02d}"
        else:
            c_id = str(c_id).strip()

        hoja_val = ws.cell(row=r, column=col_map.get("hoja", 3)).value
        hoja = str(hoja_val).strip() if hoja_val else ""

        rango_val = ws.cell(row=r, column=col_map.get("rango", 4)).value
        rango = str(rango_val).strip().upper() if rango_val else ""

        tipo_val = ws.cell(row=r, column=col_map.get("tipo_validacion", 5)).value
        tipo_str = normalizar_texto(str(tipo_val)) if tipo_val else "valor_exacto"
        # Normalizar nombres comunes
        if tipo_str in ("formula", "fórmulas", "formulas"):
            tipo_str = "formula_flexible"
        elif tipo_str in ("valor", "numero", "número", "numerico", "numérico"):
            tipo_str = "valor_numerico"
        elif tipo_str in ("formato", "format"):
            tipo_str = "formato_numero"
        elif tipo_str in ("estilo", "diseno", "diseño", "visual"):
            tipo_str = "estilo_visual_flexible"
        elif tipo_str in ("pivot", "td", "tabladinamica"):
            tipo_str = "tabla_dinamica"
        elif tipo_str in ("merge", "combinadas", "combinar"):
            tipo_str = "celdas_combinadas"
        elif tipo_str in ("validacion", "lista", "val_datos"):
            tipo_str = "validacion_datos"
        elif tipo_str in ("grafico", "graficos", "graficodinamico", "grafico_dinamico", "chart", "chartsheet", "hojagrafico", "hoja_grafico"):
            tipo_str = "grafico_dinamico"

        params_raw = ws.cell(row=r, column=col_map.get("parametros", 6)).value if "parametros" in col_map else "{}"
        params_dict = _limpiar_json_str(params_raw)

        puntos_val = ws.cell(row=r, column=col_map.get("puntos", 7)).value if "puntos" in col_map else 10.0
        try:
            puntos = float(puntos_val) if puntos_val is not None else 10.0
        except (ValueError, TypeError):
            puntos = 10.0

        ob_val = ws.cell(row=r, column=col_map.get("obligatorio", 8)).value if "obligatorio" in col_map else "NO"
        obligatorio = str(ob_val).strip().lower() in ("si", "sí", "true", "1", "yes", "s")

        criterio = CriterioRubrica(
            id=c_id,
            criterio=str(criterio_desc).strip(),
            hoja=hoja,
            rango=rango,
            tipo_validacion=tipo_str,
            parametros=params_dict,
            puntos=puntos,
            obligatorio=obligatorio
        )
        criterios.append(criterio)

    puntos_totales = sum(c.puntos for c in criterios)
    if abs(puntos_totales - 100.0) > 0.01:
        advertencias.append(
            f"La suma total de los puntos de la rúbrica es {puntos_totales:.2f} pts (se recomienda exactamente 100 pts)."
        )

    return criterios, puntos_totales, advertencias


def parsear_rubrica_desde_json(ruta_json: str) -> Tuple[List[CriterioRubrica], float, List[str]]:
    """
    Carga criterios de rúbrica desde un archivo .json.
    """
    criterios: List[CriterioRubrica] = []
    advertencias: List[str] = []

    if not os.path.isfile(ruta_json):
        return criterios, 0.0, [f"Archivo JSON no encontrado: {ruta_json}"]

    try:
        with open(ruta_json, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        return criterios, 0.0, [f"Error al leer JSON de rúbrica '{ruta_json}': {e}"]

    items = data.get("criterios", data) if isinstance(data, dict) else data
    if not isinstance(items, list):
        return criterios, 0.0, ["El contenido del archivo JSON debe ser una lista de criterios o un objeto con clave 'criterios'."]

    for idx, item in enumerate(items, 1):
        if not isinstance(item, dict):
            continue
        c_id = str(item.get("id", f"CRIT_{idx:02d}"))
        c_desc = str(item.get("criterio", item.get("descripcion", f"Criterio {idx}")))
        c_hoja = str(item.get("hoja", ""))
        c_rango = str(item.get("rango", "")).upper()
        c_tipo = normalizar_texto(str(item.get("tipo_validacion", item.get("tipo", "valor_exacto"))))
        c_params = item.get("parametros", item.get("params", {}))
        if isinstance(c_params, str):
            c_params = _limpiar_json_str(c_params)
        c_puntos = float(item.get("puntos", 10.0))
        c_ob = bool(item.get("obligatorio", False))

        criterios.append(CriterioRubrica(
            id=c_id,
            criterio=c_desc,
            hoja=c_hoja,
            rango=c_rango,
            tipo_validacion=c_tipo,
            parametros=c_params,
            puntos=c_puntos,
            obligatorio=c_ob
        ))

    puntos_totales = sum(c.puntos for c in criterios)
    if abs(puntos_totales - 100.0) > 0.01:
        advertencias.append(
            f"La suma total de los puntos de la rúbrica es {puntos_totales:.2f} pts (se recomienda exactamente 100 pts)."
        )

    return criterios, puntos_totales, advertencias


def cargar_rubrica(ruta_plantilla: str) -> Tuple[List[CriterioRubrica], float, List[str], str]:
    """
    Busca y carga la rúbrica según el orden de prioridad:
      1. Hoja '_RUBRICA' (o variante) en ruta_plantilla (.xlsx)
      2. Archivo '<plantilla>.rubrica.json'
      3. Archivo 'rubrica.json' en la misma carpeta

    Retorna:
        (criterios, puntos_totales, advertencias, origen)
    """
    advertencias: List[str] = []

    # 1. Intentar cargar desde hoja Excel de la plantilla
    if ruta_plantilla and os.path.isfile(ruta_plantilla) and ruta_plantilla.lower().endswith(".xlsx"):
        try:
            wb = openpyxl.load_workbook(ruta_plantilla, data_only=False, read_only=False)
            hoja_rubrica = None
            for name in wb.sheetnames:
                if es_rubrica(name):
                    hoja_rubrica = wb[name]
                    break

            if hoja_rubrica is not None:
                criterios, puntos, advs = parsear_rubrica_desde_hoja(hoja_rubrica)
                wb.close()
                if criterios:
                    return criterios, puntos, advs, f"Hoja Excel '{hoja_rubrica.title}' en plantilla"
                advertencias.extend(advs)
            wb.close()
        except Exception as e:
            advertencias.append(f"No se pudo inspeccionar hoja de rúbrica en '{ruta_plantilla}': {e}")

    # 2. Buscar archivo <plantilla>.rubrica.json
    if ruta_plantilla:
        base_name, _ = os.path.splitext(ruta_plantilla)
        candidato_1 = f"{base_name}{SUFIJO_ARCHIVO_RUBRICA}"
        if os.path.isfile(candidato_1):
            criterios, puntos, advs = parsear_rubrica_desde_json(candidato_1)
            if criterios:
                return criterios, puntos, advs, f"Archivo JSON '{os.path.basename(candidato_1)}'"
            advertencias.extend(advs)

        # 3. Buscar rubrica.json en la misma carpeta
        dir_plantilla = os.path.dirname(os.path.abspath(ruta_plantilla))
        candidato_2 = os.path.join(dir_plantilla, ARCHIVO_RUBRICA_GLOBAL)
        if os.path.isfile(candidato_2):
            criterios, puntos, advs = parsear_rubrica_desde_json(candidato_2)
            if criterios:
                return criterios, puntos, advs, f"Archivo JSON '{ARCHIVO_RUBRICA_GLOBAL}'"
            advertencias.extend(advs)

    return [], 0.0, advertencias + ["No se encontró ninguna rúbrica válida (ni hoja _RUBRICA ni archivo rubrica.json)."], "Ninguno"
