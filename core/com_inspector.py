# -*- coding: utf-8 -*-
# ============================================================================
# core/com_inspector.py — Inspección profunda vía COM (win32com) en Windows
# ============================================================================
"""
Módulo opcional y seguro que utiliza win32com para consultar detalles
internos de Tablas Dinámicas, Gráficos y Hojas de Gráfico de Excel.
Si win32com no está instalado o Excel no está presente, se degrada
elegantemente sin abortar la ejecución.
"""

import os
import sys
import logging
from typing import Dict, List, Tuple, Any, Optional

logger = logging.getLogger(__name__)

COM_DISPONIBLE = False
if sys.platform == "win32":
    try:
        import win32com.client
        import pythoncom
        COM_DISPONIBLE = True
    except ImportError:
        COM_DISPONIBLE = False


def iniciar_excel_com():
    """Inicia una instancia oculta de Excel vía COM."""
    if not COM_DISPONIBLE:
        return None
    try:
        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        return excel
    except Exception as e:
        logger.warning(f"No se pudo iniciar Excel vía COM: {e}")
        return None


def cerrar_excel_com(excel):
    """Cierra la instancia de Excel COM de forma segura."""
    if excel is None:
        return
    try:
        excel.Quit()
    except Exception:
        pass
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def extraer_info_pivots_com(ruta_archivo: str) -> Dict[str, List[Dict[str, Any]]]:
    """
    Inspecciona un archivo de Excel y extrae información detallada
    de todas sus Tablas Dinámicas por hoja.
    Retorna:
    {
        "NombreHoja": [
            {
                "nombre": str,
                "fuente": str,
                "campos_fila": list[str],
                "campos_columna": list[str],
                "campos_valor": list[str],
                "campos_filtro": list[str]
            }
        ]
    }
    """
    resultado: Dict[str, List[Dict[str, Any]]] = {}
    if not COM_DISPONIBLE or not os.path.isfile(ruta_archivo):
        return resultado

    excel = None
    try:
        excel = iniciar_excel_com()
        if excel is None:
            return resultado

        wb = excel.Workbooks.Open(os.path.abspath(ruta_archivo), ReadOnly=True)
        try:
            for i in range(1, wb.Worksheets.Count + 1):
                ws = wb.Worksheets(i)
                nombre_hoja = ws.Name
                pt_count = ws.PivotTables().Count
                if pt_count == 0:
                    continue

                lista_pts = []
                for j in range(1, pt_count + 1):
                    pt = ws.PivotTables(j)
                    info = {
                        "nombre": str(pt.Name),
                        "fuente": str(pt.SourceData if hasattr(pt, 'SourceData') else ""),
                        "campos_fila": [pt.RowFields(f).Name for f in range(1, pt.RowFields.Count + 1)] if hasattr(pt, 'RowFields') else [],
                        "campos_columna": [pt.ColumnFields(f).Name for f in range(1, pt.ColumnFields.Count + 1)] if hasattr(pt, 'ColumnFields') else [],
                        "campos_valor": [pt.DataFields(f).Name for f in range(1, pt.DataFields.Count + 1)] if hasattr(pt, 'DataFields') else [],
                        "campos_filtro": [pt.PageFields(f).Name for f in range(1, pt.PageFields.Count + 1)] if hasattr(pt, 'PageFields') else [],
                    }
                    lista_pts.append(info)

                if lista_pts:
                    resultado[nombre_hoja] = lista_pts
        finally:
            wb.Close(SaveChanges=False)
    except Exception as e:
        logger.warning(f"Error al inspeccionar Tablas Dinámicas COM en '{ruta_archivo}': {e}")
    finally:
        cerrar_excel_com(excel)

    return resultado


def comparar_pivots_y_graficos_com(
    ruta_plantilla: str,
    ruta_estudiante: str,
    mapa_hojas: Optional[Dict[str, Dict[str, Any]]] = None
) -> Tuple[int, List[str]]:
    """
    Compara Tablas Dinámicas y Hojas de Gráfico entre plantilla y estudiante (Modo A).
    """
    if not COM_DISPONIBLE:
        return 0, []

    errores = 0
    detalles = []
    excel = None

    try:
        excel = iniciar_excel_com()
        if excel is None:
            return 0, []

        wb_p = excel.Workbooks.Open(os.path.abspath(ruta_plantilla), ReadOnly=True)
        wb_e = excel.Workbooks.Open(os.path.abspath(ruta_estudiante), ReadOnly=True)

        try:
            # 1. Comparar Hojas de Gráficos (Chart Sheets)
            charts_p = [wb_p.Charts(k).Name for k in range(1, wb_p.Charts.Count + 1)]
            charts_e = [wb_e.Charts(k).Name for k in range(1, wb_e.Charts.Count + 1)]

            for c_nom in charts_p:
                c_nom_mapped = mapa_hojas.get(c_nom, {}).get("hoja_estudiante") if mapa_hojas else None
                if c_nom not in charts_e and (not c_nom_mapped or c_nom_mapped not in charts_e):
                    detalles.append(f"❌ Gráfico Dinámico faltante: hoja de gráfico '{c_nom}' no encontrada.")
                    errores += 1

            # 2. Comparar Tablas Dinámicas
            for i in range(1, wb_p.Worksheets.Count + 1):
                ws_p = wb_p.Worksheets(i)
                nom_p = ws_p.Name
                if nom_p.lower().startswith("_rub"):
                    continue

                nom_e = nom_p
                if mapa_hojas and nom_p in mapa_hojas:
                    nom_e = mapa_hojas[nom_p].get("hoja_estudiante") or nom_p

                try:
                    ws_e = wb_e.Worksheets(nom_e)
                except Exception:
                    continue

                pt_p_count = ws_p.PivotTables().Count
                pt_e_count = ws_e.PivotTables().Count

                if pt_p_count > 0 and pt_e_count == 0:
                    detalles.append(f"❌ Hoja '{nom_p}': Se esperaban {pt_p_count} Tablas Dinámicas pero no se encontró ninguna.")
                    errores += pt_p_count
                elif pt_p_count != pt_e_count:
                    detalles.append(f"❌ Hoja '{nom_p}': Cantidad de Tablas Dinámicas esperadas: {pt_p_count} | Encontradas: {pt_e_count}")
                    errores += abs(pt_p_count - pt_e_count)

        finally:
            wb_p.Close(SaveChanges=False)
            wb_e.Close(SaveChanges=False)

    except Exception as e:
        detalles.append(f"⚠️ Alerta en inspección COM: {e}")
    finally:
        cerrar_excel_com(excel)

    return errores, detalles
