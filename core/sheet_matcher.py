# -*- coding: utf-8 -*-
# ============================================================================
# core/sheet_matcher.py — Emparejamiento inteligente de hojas de cálculo
# ============================================================================
"""
Realiza la vinculación flexible entre las hojas de la plantilla y las
hojas creadas o renombradas por los estudiantes mediante un algoritmo
multinivel (coincidencia exacta, normalizada, por número de actividad,
similitud difusa de Levenshtein/difflib y huella de encabezados).
"""

import re
import unicodedata
from difflib import SequenceMatcher
from typing import List, Dict, Optional, Any
from core.config import NOMBRES_HOJA_RUBRICA


def normalizar_texto(texto: str) -> str:
    """Normaliza un texto: minúsculas, sin tildes, sin espacios repetidos."""
    if not texto:
        return ""
    texto_norm = unicodedata.normalize("NFKD", texto)
    texto_sin_tildes = "".join(c for c in texto_norm if not unicodedata.combining(c))
    texto_limpio = re.sub(r"\s+", " ", texto_sin_tildes).strip().lower()
    return texto_limpio


def es_rubrica(nombre_hoja: str) -> bool:
    """Verifica si una hoja corresponde a la hoja de rúbrica."""
    if not nombre_hoja:
        return False
    norm = normalizar_texto(nombre_hoja)
    return norm in NOMBRES_HOJA_RUBRICA or norm.startswith("_rub") or norm == "rubrica"


def extraer_numero_actividad(nombre: str) -> Optional[int]:
    """
    Extrae el número identificador de la actividad al inicio del nombre.
    Ejemplos:
        '1. Ventas Semanales' -> 1
        'Actividad 3 - Fórmulas' -> 3
        'Hoja 2' -> 2
    """
    if not nombre:
        return None
    m = re.match(r"^\s*(?:hoja|actividad|ejercicio|p|part|taller)?\s*(\d+)", nombre, re.IGNORECASE)
    if m:
        try:
            return int(m.group(1))
        except ValueError:
            return None
    return None


def obtener_huella_encabezados(ws) -> set:
    """
    Extrae palabras clave de los encabezados (primeras 5 filas y 15 columnas).
    """
    palabras = set()
    if ws is None:
        return palabras
    try:
        max_r = min(ws.max_row or 5, 5)
        max_c = min(ws.max_column or 15, 15)
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                val = ws.cell(row=r, column=c).value
                if val and isinstance(val, str):
                    norm = normalizar_texto(val)
                    for token in norm.split():
                        if len(token) >= 3:
                            palabras.add(token)
    except Exception:
        pass
    return palabras


def emparejar_hojas(
    hojas_plantilla: List[str],
    hojas_estudiante: List[str],
    wb_plantilla=None,
    wb_estudiante=None
) -> Dict[str, Dict[str, Any]]:
    """
    Realiza un emparejamiento multinivel entre hojas requeridas y hojas del estudiante.

    Retorna:
    {
        nombre_plantilla: {
            "hoja_estudiante": str | None,
            "metodo": str,
            "score": float
        }
    }
    """
    disponibles = [h for h in hojas_estudiante if not es_rubrica(h)]
    mapa: Dict[str, Dict[str, Any]] = {}
    plantilla_req = [h for h in hojas_plantilla if not es_rubrica(h)]

    # 1. Coincidencia Exacta
    for p in plantilla_req:
        if p in disponibles:
            mapa[p] = {
                "hoja_estudiante": p,
                "metodo": "exacto",
                "score": 1.0
            }
            disponibles.remove(p)

    # 2. Coincidencia Normalizada (sin tildes ni mayúsculas)
    for p in plantilla_req:
        if p in mapa:
            continue
        p_norm = normalizar_texto(p)
        match_e = None
        for e in disponibles:
            if normalizar_texto(e) == p_norm:
                match_e = e
                break
        if match_e:
            mapa[p] = {
                "hoja_estudiante": match_e,
                "metodo": "normalizado",
                "score": 0.95
            }
            disponibles.remove(match_e)

    # 3. Coincidencia por Número de Actividad
    for p in plantilla_req:
        if p in mapa:
            continue
        num_p = extraer_numero_actividad(p)
        if num_p is not None:
            candidatos = [e for e in disponibles if extraer_numero_actividad(e) == num_p]
            if len(candidatos) == 1:
                e_sel = candidatos[0]
                mapa[p] = {
                    "hoja_estudiante": e_sel,
                    "metodo": f"prefijo_numérico ({num_p})",
                    "score": 0.90
                }
                disponibles.remove(e_sel)
            elif len(candidatos) > 1:
                p_norm = normalizar_texto(p)
                mejor_e = max(candidatos, key=lambda c: SequenceMatcher(None, p_norm, normalizar_texto(c)).ratio())
                mapa[p] = {
                    "hoja_estudiante": mejor_e,
                    "metodo": f"prefijo_numérico_desempate ({num_p})",
                    "score": 0.85
                }
                disponibles.remove(mejor_e)

    # 4. Similitud Textual Difusa (Fuzzy Matching)
    for p in plantilla_req:
        if p in mapa:
            continue
        p_norm = normalizar_texto(p)
        mejor_e = None
        mejor_score = 0.0
        for e in disponibles:
            e_norm = normalizar_texto(e)
            ratio = SequenceMatcher(None, p_norm, e_norm).ratio()
            if p_norm and e_norm and (p_norm in e_norm or e_norm in p_norm):
                ratio = max(ratio, 0.75)
            if ratio > mejor_score and ratio >= 0.60:
                mejor_score = ratio
                mejor_e = e
        if mejor_e:
            mapa[p] = {
                "hoja_estudiante": mejor_e,
                "metodo": f"fuzzy ({mejor_score:.2f})",
                "score": mejor_score
            }
            disponibles.remove(mejor_e)

    # 5. Huella de Contenido (Encabezados)
    if wb_plantilla is not None and wb_estudiante is not None:
        for p in plantilla_req:
            if p in mapa:
                continue
            try:
                if p in wb_plantilla.sheetnames:
                    huella_p = obtener_huella_encabezados(wb_plantilla[p])
                    if len(huella_p) >= 3:
                        mejor_e = None
                        mejor_jaccard = 0.0
                        for e in disponibles:
                            if e in wb_estudiante.sheetnames:
                                huella_e = obtener_huella_encabezados(wb_estudiante[e])
                                if huella_e:
                                    inter = len(huella_p.intersection(huella_e))
                                    union = len(huella_p.union(huella_e))
                                    jaccard = inter / union if union > 0 else 0
                                    if jaccard > mejor_jaccard and jaccard >= 0.35:
                                        mejor_jaccard = jaccard
                                        mejor_e = e
                        if mejor_e:
                            mapa[p] = {
                                "hoja_estudiante": mejor_e,
                                "metodo": f"contenido ({mejor_jaccard:.2f})",
                                "score": mejor_jaccard
                            }
                            disponibles.remove(mejor_e)
            except Exception:
                pass

    # 6. Emparejamiento de Hojas de Gráfico (ChartSheets) restantes
    if wb_plantilla is not None and wb_estudiante is not None:
        for p in plantilla_req:
            if p in mapa:
                continue
            ws_p = wb_plantilla[p] if p in wb_plantilla.sheetnames else None
            es_cs_p = ws_p is not None and not hasattr(ws_p, "max_row")
            if es_cs_p or "grafico" in normalizar_texto(p):
                candidatos_cs = []
                for e in disponibles:
                    ws_e = wb_estudiante[e] if e in wb_estudiante.sheetnames else None
                    es_cs_e = ws_e is not None and not hasattr(ws_e, "max_row")
                    if es_cs_e or "grafico" in normalizar_texto(e):
                        candidatos_cs.append(e)
                if candidatos_cs:
                    e_sel = candidatos_cs[0]
                    mapa[p] = {
                        "hoja_estudiante": e_sel,
                        "metodo": "chartsheet_fallback",
                        "score": 0.75
                    }
                    disponibles.remove(e_sel)

    # Hojas no encontradas
    for p in plantilla_req:
        if p not in mapa:
            mapa[p] = {
                "hoja_estudiante": None,
                "metodo": "no_encontrada",
                "score": 0.0
            }

    return mapa
