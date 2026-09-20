# -*- coding: utf-8 -*-
# ============================================================================
# core/validators/formula_validator.py — Validador flexible de fórmulas
# ============================================================================
"""
Validador para 'formula_flexible'.
Evalúa la presencia de fórmulas, funciones requeridas (con traducción
automática Español/Inglés) y referencias clave a rangos sin exigir
sintaxis textual idéntica.
"""

import re
from typing import List, Set, Any
from core.validators.base import BaseValidator
from core.models import CriterioRubrica, ResultadoCriterio
from core.config import FUNCIONES_ES_A_EN, FUNCIONES_EN_A_ES


def extraer_funciones_formula(formula: str) -> Set[str]:
    """
    Extrae los nombres de las funciones utilizadas en una fórmula.
    Normaliza prefijos internos como _xlfn. y retorna los nombres en mayúsculas.
    """
    if not formula or not isinstance(formula, str):
        return set()

    limpia = formula.strip().upper()
    limpia = limpia.replace("_XLFN._XLWS.", "").replace("_XLFN.", "").replace("=@", "=").replace("(@", "(")

    # Coincidir palabras seguidas de paréntesis de apertura: NOMBRE(
    patron = r'([A-Z0-9_\.]+)\s*\('
    encontradas = re.findall(patron, limpia)
    resultado = set()
    for f in encontradas:
        f_upper = f.upper()
        resultado.add(f_upper)
        # Agregar equivalente en inglés y español si existe
        if f_upper in FUNCIONES_ES_A_EN:
            resultado.add(FUNCIONES_ES_A_EN[f_upper])
        if f_upper in FUNCIONES_EN_A_ES:
            resultado.add(FUNCIONES_EN_A_ES[f_upper])

    return resultado


def normalizar_referencia(ref: str) -> str:
    """Limpia una referencia eliminando '$' y espacios."""
    if not ref:
        return ""
    return ref.replace("$", "").replace(" ", "").upper()


class FormulaValidator(BaseValidator):
    """Valida fórmulas de forma semántica y tolerante."""

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

        # Funciones requeridas
        funcs_req_param = params.get("funciones", params.get("funcion", []))
        if isinstance(funcs_req_param, str):
            funcs_req_param = [funcs_req_param]

        # Normalizar funciones requeridas a un conjunto canónico (inglés y español)
        funcs_requeridas_canonicas: Set[str] = set()
        for f in funcs_req_param:
            f_u = f.strip().upper()
            funcs_requeridas_canonicas.add(f_u)
            if f_u in FUNCIONES_ES_A_EN:
                funcs_requeridas_canonicas.add(FUNCIONES_ES_A_EN[f_u])
            if f_u in FUNCIONES_EN_A_ES:
                funcs_requeridas_canonicas.add(FUNCIONES_EN_A_ES[f_u])

        # Funciones prohibidas
        funcs_prohibidas_param = params.get("prohibidas", params.get("prohibida", []))
        if isinstance(funcs_prohibidas_param, str):
            funcs_prohibidas_param = [funcs_prohibidas_param]
        funcs_prohibidas = {f.strip().upper() for f in funcs_prohibidas_param}

        # Referencias obligatorias
        refs_req = params.get("referencias", params.get("referencia", []))
        if isinstance(refs_req, str):
            refs_req = [refs_req]
        refs_requeridas_norm = [normalizar_referencia(r) for r in refs_req if r]

        fallos: List[str] = []
        celdas_erroneas: List[str] = []

        for coord in coords:
            celda = ws_estudiante[coord] if ws_estudiante else None
            valor = celda.value if celda is not None else None

            # 1. Verificar si es fórmula
            if not isinstance(valor, str) or not valor.strip().startswith("="):
                celdas_erroneas.append(coord)
                fallos.append(f"Celda {coord}: no contiene fórmula (valor: '{valor}')")
                continue

            formula_raw = valor.strip()
            formula_norm = normalizar_referencia(formula_raw)
            funcs_encontradas = extraer_funciones_formula(formula_raw)

            # 2. Verificar funciones requeridas
            if funcs_requeridas_canonicas:
                coincide_alguna = bool(funcs_requeridas_canonicas.intersection(funcs_encontradas))
                if not coincide_alguna:
                    celdas_erroneas.append(coord)
                    esperadas_legibles = "/".join(sorted(funcs_requeridas_canonicas))
                    fallos.append(f"Celda {coord}: no utiliza la función requerida ({esperadas_legibles}). Fórmula: '{formula_raw}'")
                    continue

            # 3. Verificar funciones prohibidas
            if funcs_prohibidas:
                encontradas_prohibidas = funcs_prohibidas.intersection(funcs_encontradas)
                if encontradas_prohibidas:
                    celdas_erroneas.append(coord)
                    fallos.append(f"Celda {coord}: utiliza función no permitida ({', '.join(encontradas_prohibidas)})")
                    continue

            # 4. Verificar referencias
            if refs_requeridas_norm:
                refs_faltantes = []
                for ref_esperada in refs_requeridas_norm:
                    if ref_esperada not in formula_norm:
                        refs_faltantes.append(ref_esperada)
                if refs_faltantes:
                    celdas_erroneas.append(coord)
                    fallos.append(f"Celda {coord}: no referencia el rango/celda '{', '.join(refs_faltantes)}'")
                    continue

            # 5. Verificar fragmentos, operadores o textos requeridos (ej: "JEFE", "5%/0.05", etc.)
            contiene_req = params.get("contiene", params.get("fragmentos", []))
            if isinstance(contiene_req, str):
                contiene_req = [contiene_req]

            faltantes_contiene = []
            formula_limpia = formula_raw.upper().replace(" ", "")
            for elemento in contiene_req:
                if isinstance(elemento, (list, tuple)):
                    opciones = [str(o).strip().upper().replace(" ", "") for o in elemento]
                elif "/" in str(elemento):
                    opciones = [str(o).strip().upper().replace(" ", "") for o in str(elemento).split("/")]
                else:
                    opciones = [str(elemento).strip().upper().replace(" ", "")]

                if not any(opt in formula_limpia for opt in opciones):
                    faltantes_contiene.append(str(elemento))

            if faltantes_contiene:
                celdas_erroneas.append(coord)
                fallos.append(f"Celda {coord}: la fórmula no contiene el cálculo o elemento requerido ({', '.join(faltantes_contiene)})")
                continue

        aprobado = len(fallos) == 0
        puntos = criterio.puntos if aprobado else 0.0
        mensaje = "Fórmula correcta." if aprobado else "; ".join(fallos[:3])
        if len(fallos) > 3:
            mensaje += f" (y {len(fallos)-3} celdas más con errores en fórmula)"

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
