# -*- coding: utf-8 -*-
# ============================================================================
# core/config.py — Configuración central de APP_REVISA_EXCEL_V2
# ============================================================================
"""
Define constantes, rutas predeterminadas, paletas cromáticas,
mensajes del sistema y diccionarios de traducción de fórmulas para
el motor híbrido de evaluación (Modo A y Modo B).
"""

import os
from typing import Dict, List, Tuple

# --- Rutas del Proyecto ---
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PLANTILLAS_DIR = os.path.join(BASE_DIR, "PLANTILLAS")
DEFAULT_PLANTILLA_PATH = os.path.join(PLANTILLAS_DIR, "PLANTILLA.xlsx")
TRABAJOS_DIR = os.path.join(BASE_DIR, "TRABAJOS_ESTUDIANTES")
REVISADOS_DIR = os.path.join(BASE_DIR, "REVISADOS")
LOGS_DIR = os.path.join(BASE_DIR, "LOGS")

SUFIJO_REVISADO = "_rev"
EXTENSIONES_VALIDAS = (".xlsx",)

# --- Nombres especiales de hojas y rúbricas ---
HOJA_RUBRICA_DEFAULT = "_RUBRICA"
NOMBRES_HOJA_RUBRICA = ("_rubrica", "rubrica", "rúbrica", "_rubricas", "rubricas")
SUFIJO_ARCHIVO_RUBRICA = ".rubrica.json"
ARCHIVO_RUBRICA_GLOBAL = "rubrica.json"

# --- Colores de Retroalimentación Visual (ARGB Hex para openpyxl) ---
COLOR_ERROR_FILL = "FFFFCCCC"        # Rojo suave / pastel (#FFCCCC)
COLOR_ERROR_FONT = "FF9C0006"        # Rojo oscuro para texto
COLOR_SUCCESS_FILL = "FFE2EFDA"      # Verde suave / pastel (#E2EFDA)
COLOR_SUCCESS_FONT = "FF276A3C"      # Verde oscuro para texto
COLOR_BORDER_GRID = "FFCCCCCC"       # Gris suave de rejilla
COLOR_HEADER_BG = "FF1A365D"         # Azul corporativo profundo (#1a365d)
COLOR_HEADER_FG = "FFFFFFFF"         # Blanco

# --- Patrones de Nombres Genéricos / Por Defecto de Excel ---
NOMBRES_TABLA_DEFECTO = [
    r"tabla\s*\d+",
    r"table\s*\d+",
]

NOMBRES_HOJA_DEFECTO = [
    r"hoja\s*\d+",
    r"sheet\s*\d+",
    r"hoja\s*\d+\s*\(\d+\)",
    r"sheet\s*\d+\s*\(\d+\)",
]

NOMBRES_GRAFICO_HOJA_DEFECTO = [
    r"gr[aá]fico\s*\d+",
    r"chart\s*\d+",
    r"graph\s*\d+",
]

NOMBRES_TD_DEFECTO = [
    r"tabla\s*din[aá]mica\s*\d+",
    r"pivot\s*table\s*\d+",
]

# --- Traducción de Funciones de Excel (Español <-> Inglés) ---
# En archivos .xlsx las fórmulas siempre se guardan internamente en inglés.
FUNCIONES_ES_A_EN: Dict[str, str] = {
    "SUMA": "SUM",
    "PROMEDIO": "AVERAGE",
    "CONTAR": "COUNT",
    "CONTARA": "COUNTA",
    "CONTAR.SI": "COUNTIF",
    "CONTAR.SI.CONJUNTO": "COUNTIFS",
    "SUMAR.SI": "SUMIF",
    "SUMAR.SI.CONJUNTO": "SUMIFS",
    "PROMEDIO.SI": "AVERAGEIF",
    "PROMEDIO.SI.CONJUNTO": "AVERAGEIFS",
    "BUSCARV": "VLOOKUP",
    "BUSCARH": "HLOOKUP",
    "BUSCARX": "XLOOKUP",
    "COINCIDIR": "MATCH",
    "COINCIDIRX": "XMATCH",
    "INDICE": "INDEX",
    "SI": "IF",
    "SI.ERROR": "IFERROR",
    "SI.CONJUNTO": "IFS",
    "Y": "AND",
    "O": "OR",
    "NO": "NOT",
    "VERDADERO": "TRUE",
    "FALSO": "FALSE",
    "CONCATENAR": "CONCATENATE",
    "CONCAT": "CONCAT",
    "UNIRCADENAS": "TEXTJOIN",
    "TEXTO": "TEXT",
    "HOY": "TODAY",
    "AHORA": "NOW",
    "AÑO": "YEAR",
    "MES": "MONTH",
    "DIA": "DAY",
    "DIAS": "DAYS",
    "FECHA": "DATE",
    "REDONDEAR": "ROUND",
    "REDONDEAR.MAS": "ROUNDUP",
    "REDONDEAR.MENOS": "ROUNDDOWN",
    "ENTERO": "INT",
    "POTENCIA": "POWER",
    "RAIZ": "SQRT",
    "ABS": "ABS",
    "MAX": "MAX",
    "MIN": "MIN",
    "MAYUSC": "UPPER",
    "MINUSC": "LOWER",
    "NOMPROPIO": "PROPER",
    "LARGO": "LEN",
    "IZQUIERDA": "LEFT",
    "DERECHA": "RIGHT",
    "EXTRAE": "MID",
    "ENCONTRAR": "FIND",
    "HALLAR": "SEARCH",
    "SUSTITUIR": "SUBSTITUTE",
    "REEMPLAZAR": "REPLACE",
    "ESPACIOS": "TRIM",
    "DESREF": "OFFSET",
    "FILA": "ROW",
    "COLUMNA": "COLUMN",
    "TRANSPONER": "TRANSPOSE",
    "ELEGIR": "CHOOSE",
    "ALEATORIO": "RAND",
    "ALEATORIO.ENTRE": "RANDBETWEEN",
    "RESIDUO": "MOD",
    "PRODUCTO": "PRODUCT",
    "MEDIANA": "MEDIAN",
    "MODA": "MODE",
    "K.ESIMO.MAYOR": "LARGE",
    "K.ESIMO.MENOR": "SMALL",
    "ORDENAR": "SORT",
    "FILTRAR": "FILTER",
    "UNICOS": "UNIQUE",
    "UNICO": "UNIQUE",
    "TIPO": "TYPE",
    "ESERROR": "ISERROR",
    "ESNUMERO": "ISNUMBER",
    "ESTEXTO": "ISTEXT",
    "ESBLANCO": "ISBLANK",
}

FUNCIONES_EN_A_ES: Dict[str, str] = {v: k for k, v in FUNCIONES_ES_A_EN.items()}

# Funciones de agregación donde los argumentos son conmutativos
FUNCIONES_AGREGACION_CONMUTATIVAS = {
    "SUM", "AVERAGE", "MIN", "MAX", "COUNT", "COUNTA", "PRODUCT", "MEDIAN"
}

# --- Catálogo de Validadores Disponibles en Modo B ---
VALIDADORES_DISPONIBLES = [
    "valor_exacto",
    "valor_numerico",
    "formula_flexible",
    "formato_numero",
    "estilo_visual_flexible",
    "tabla_dinamica",
    "celdas_combinadas",
    "validacion_datos",
]

# --- Mensajes Estándar ---
MENSAJES = {
    "valor": "❌ Valor incorrecto. Esperado: '{esperado}' | Encontrado: '{encontrado}'",
    "formula": "❌ Fórmula incorrecta. Esperada: '{esperado}' | Encontrada: '{encontrado}'",
    "funcion": "❌ Función utilizada incorrecta. Esperada: '{esperado}' | Encontrada: '{encontrado}'",
    "referencia_faltante": "❌ Referencia obligatoria faltante en fórmula: '{ref}'",
    "formato_num": "❌ Formato numérico incorrecto. Esperado tipo: '{esperado}' | Encontrado formato: '{encontrado}'",
    "estilo": "❌ Atributo visual incorrecto ({atributo}). Esperado: {esperado} | Encontrado: {encontrado}",
    "tabla_faltante": "❌ Tabla de Excel faltante: '{nombre}'",
    "validacion_faltante": "❌ Validación de datos no encontrada en rango '{rango}'",
    "merge_faltante": "❌ El rango '{rango}' debe estar combinado y no lo está.",
    "td_faltante": "❌ Tabla Dinámica faltante en la hoja '{hoja}'.",
    "td_campo_faltante": "❌ Campo '{campo}' no encontrado en {seccion} de la Tabla Dinámica.",
    "desplazamiento_detectado": "⚠️ ADVERTENCIA: Se detectó un posible desplazamiento de {filas} fila(s) y {columnas} col(s) en la hoja '{hoja}'.",
}
