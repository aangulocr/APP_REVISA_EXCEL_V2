# APP_REVISA_EXCEL_V2 — Motor Híbrido de Auditoría y Evaluación de Excel

Sistema autónomo, modular y 100% offline para la revisión, auditoría y calificación inteligente de hojas de cálculo de Microsoft Excel (`.xlsx`). Diseñado para docentes, evaluadores y administradores educativos.

---

## 🚀 Novedades Clave de la Versión 2.0 (Arquitectura Híbrida)

La versión 2.0 evoluciona el sistema hacia una **arquitectura híbrida de evaluación**:

1. **MODO B (Rúbrica Inteligente por Criterios - Smart Rubric) [NUEVO & RECOMENDADO]:**
   - Evaluación declarativa orientada a competencias ponderada de **0 a 100 puntos**.
   - Reglas configurables directamente por el profesor desde una hoja de cálculo oculta o visible llamada `_RUBRICA` dentro del archivo de plantilla, o mediante un archivo `rubrica.json`.
   - Catálogo de validadores inteligentes con soporte para **tolerancias numéricas absolutas/porcentuales**, **fórmulas bilingües (Español/Inglés)**, **formatos numéricos agnósticos a configuración regional**, **estilos visuales flexibles**, **celdas combinadas**, **validación de datos** y **tablas dinámicas**.
   - Retroalimentación visual directa en el archivo del estudiante (`_rev.xlsx`): resaltado en rojo pastel (`#FFCCCC`) con comentarios explicativos detallados (`❌ [FALLÓ: -X pts]`) y marcas de aprobación (`✅ [APROBADO: +X pts]`).

2. **MODO A (Plantilla Espejo - Strict Diff Mejorado):**
   - Comparación celda a celda y objeto a objeto contra plantilla idéntica para ejercicios que exigen una réplica exacta.
   - **Detector Heurístico de Desplazamientos:** Analiza si un archivo con alta tasa de fallos está simplemente desplazado por haber insertado o borrado una fila o columna adicional (ej. `+1 fila abajo`), alertando al docente en el reporte sin falsos diagnósticos.

3. **Interfaz Gráfica Web Moderna (100% Offline):**
   - Panel de control visual con diseño **Glassmorphism**, modo oscuro/claro y tipografía del sistema (cero dependencias de internet o CDNs externos).
   - **Pre-chequeo e Inspector de Rúbrica:** Muestra en tiempo real la lista de criterios y valida si suman exactamente **100.0 puntos** antes de iniciar la auditoría.
   - **Generador de Plantillas de Rúbrica:** Crea con un clic un archivo Excel de ejemplo (`PLANTILLA_RUBRICA_EJEMPLO.xlsx`) con datos reales, fórmulas y la hoja `_RUBRICA` prellenada a 100 puntos.
   - **Consola en Vivo con Barra de Progreso:** Transmisión de logs en tiempo real vía Server-Sent Events (SSE) con indicador interactivo de porcentaje y estudiante evaluado.

---

## 📂 Estructura del Proyecto

```
APP_REVISA_EXCEL_V2/
├── core/
│   ├── config.py                 # Constantes, rutas, mensajes, colores, traducción ES/EN
│   ├── models.py                 # Dataclasses tipadas (CriterioRubrica, ResultadoEstudiante)
│   ├── rubric_parser.py          # Lector y validador de _RUBRICA (Excel) o rubrica.json
│   ├── rubric_template.py        # Generador de plantilla Excel modelo con rúbrica
│   ├── sheet_matcher.py          # Emparejamiento inteligente multinivel de hojas
│   ├── shift_detector.py         # Detector heurístico de filas/columnas desplazadas
│   ├── validators/               # Catálogo de validadores del Modo B
│   │   ├── base.py               # Clase base abstracta BaseValidator
│   │   ├── value_validator.py    # valor_exacto, valor_numerico con tolerancias
│   │   ├── formula_validator.py  # formula_flexible con traducción ES/EN y referencias
│   │   ├── format_validator.py   # formato_numero (moneda, porcentaje, fecha, decimal)
│   │   ├── style_validator.py    # estilo_visual_flexible (negrita, relleno, bordes)
│   │   ├── structure_validator.py# celdas_combinadas y validacion_datos
│   │   └── pivot_validator.py    # tabla_dinamica
│   ├── evaluators/               # Motores de ejecución
│   │   ├── base_evaluator.py     # Utilidades de marcado de celdas y comentarios
│   │   ├── mirror_evaluator.py   # Modo A: Plantilla espejo + shift detector
│   │   └── rubric_evaluator.py   # Modo B: Rúbrica ponderada 0-100 pts
│   ├── com_inspector.py          # Inspección COM opcional protegida (win32com)
│   └── report_generator.py       # Generador de REPORTE_NOTAS.csv, JSON y LOG_NOTAS.xlsx
├── gui/
│   ├── index.html                # Interfaz web Glassmorphic 100% offline
│   ├── style.css                 # Estilos modernos con paleta HSL y responsive
│   └── app.js                    # Controlador de cliente y streaming SSE
├── PLANTILLAS/                   # Plantillas maestras de evaluación
├── TRABAJOS_ESTUDIANTES/         # Carpeta de entrada con trabajos a evaluar (.xlsx)
├── REVISADOS/                    # Carpeta de salida con archivos calificados (_rev.xlsx)
├── LOGS/                         # Carpeta de salida de reportes y logs
├── tests/                        # Suite de 16 pruebas automatizadas con pytest
├── auditor.py                    # Fachada orquestadora
├── main.py                       # CLI ejecutable con argumentos argparse
├── gui_server.py                 # Servidor HTTP local con SSE y API REST
├── crear_datos_prueba.py         # Generador de casos de prueba
├── requirements.txt              # openpyxl y pywin32
├── EJECUTAR_INTERFAZ.bat         # Lanzador GUI en Windows (Doble clic)
├── EJECUTAR_AUDITORIA.bat        # Lanzador CLI en Windows (Doble clic)
├── EJECUTAR_INTERFAZ.sh          # Lanzador GUI en Linux / macOS
└── EJECUTAR_AUDITORIA.sh         # Lanzador CLI en Linux / macOS
```

---

## 🎯 Guía de Configuración de la Rúbrica (Modo B)

### Estructura de la Hoja `_RUBRICA` en Excel
Crea o edita una hoja llamada `_RUBRICA` (puede ser visible u oculta) en el archivo de plantilla `.xlsx` con las siguientes columnas en la fila 1:

| ID | CRITERIO | HOJA | RANGO | TIPO_VALIDACION | PARAMETROS | PUNTOS | OBLIGATORIO |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| `CRIT_01` | Calcular Subtotal multiplicando Cantidad por Precio | `Ventas` | `E2:E10` | `formula_flexible` | `{"referencias": ["C", "D"]}` | `15` | `SI` |
| `CRIT_02` | Calcular Total con función SUMA | `Ventas` | `G11` | `formula_flexible` | `{"funciones": ["SUMA", "SUM"]}` | `20` | `SI` |
| `CRIT_03` | Gran Total de Ventas esperado | `Ventas` | `G11` | `valor_numerico` | `{"valor": 5006.5, "tolerancia": 1.0}` | `15` | `NO` |
| `CRIT_04` | Formato de moneda en precios y totales | `Ventas` | `D2:G11` | `formato_numero` | `{"formato": "moneda"}` | `10` | `NO` |
| `CRIT_05` | Estilo de encabezados (negrita y relleno) | `Ventas` | `A1:G1` | `estilo_visual_flexible` | `{"requiere_negrita": true, "requiere_relleno": true}` | `10` | `NO` |
| `CRIT_06` | Título combinado en hoja Resumen | `Resumen` | `A1:D1` | `celdas_combinadas` | `{}` | `10` | `NO` |
| `CRIT_07` | Lista desplegable en celda B3 | `Resumen` | `B3` | `validacion_datos` | `{"tipo": "list"}` | `10` | `NO` |
| `CRIT_08` | Tabla Dinámica de ventas | `Resumen` | `A10` | `tabla_dinamica` | `{"campos_fila": ["Categoria"]}` | `10` | `NO` |

### Catálogo de Validadores y Parámetros

1. **`valor_exacto` / `valor_numerico`:**
   - Parámetros: `{"valor": 1500, "tolerancia": 0.05, "tolerancia_pct": 0.02}`
   - Para texto: normalización automática sin tildes, minúsculas y sin espacios duplicados (ej: `"México"` coincide con `"mexico"`).
2. **`formula_flexible`:**
   - Parámetros: `{"funciones": ["SUMA", "SUM"], "referencias": ["C2:C10"], "prohibidas": ["IF"]}`
   - Reconoce equivalencias conmutativas (ej: `A+B` == `B+A`, `SUMA` == `SUM`).
3. **`formato_numero`:**
   - Parámetros: `{"formato": "moneda"}` (opciones: `moneda`, `porcentaje`, `fecha`, `decimal`, `entero`).
   - Agnóstico al símbolo regional (`$`, `€`, `₡`, etc.).
4. **`estilo_visual_flexible`:**
   - Parámetros: `{"requiere_negrita": true, "requiere_relleno": true, "requiere_bordes": true, "alineacion": "center"}`.
5. **`celdas_combinadas`:**
   - Verifica que el rango especificado esté combinado (`merged`).
6. **`validacion_datos`:**
   - Parámetros: `{"tipo": "list"}` (o `whole`, `decimal`, etc.).
7. **`tabla_dinamica`:**
   - Parámetros: `{"campos_fila": ["Sucursal"], "campos_valor": ["Ventas"]}`.

---

## 💻 Instrucciones de Uso

### 1. Panel de Control Gráfico Web (Recomendado)
Haz doble clic en:
```cmd
EJECUTAR_INTERFAZ.bat
```
*Se abrirá automáticamente el navegador en `http://localhost:5000`.*
1. Elige el **Modo de Evaluación** (Modo B Rúbrica o Modo A Espejo).
2. Selecciona la plantilla y la carpeta con los archivos de los estudiantes.
3. Si estás en Modo B, el panel mostrará de inmediato la **tarjeta de pre-chequeo de rúbrica** confirmando el total de puntos.
4. Presiona **Iniciar Auditoría** y observa la barra de progreso y la consola en tiempo real.

### 2. Ejecución por Línea de Comandos (CLI)
Haz doble clic en `EJECUTAR_AUDITORIA.bat` o ejecuta:
```powershell
# Modo Automático (detecta si hay rúbrica en la plantilla)
.\.venv\Scripts\python.exe main.py

# Forzar Modo B (Rúbrica Inteligente)
.\.venv\Scripts\python.exe main.py --modo modo_b_rubrica --plantilla "PLANTILLAS/PLANTILLA.xlsx" --trabajos "TRABAJOS_ESTUDIANTES"

# Forzar Modo A (Plantilla Espejo con detector de desplazamientos)
.\.venv\Scripts\python.exe main.py --modo modo_a_espejo

# Generar plantilla de rúbrica de ejemplo (.xlsx)
.\.venv\Scripts\python.exe main.py --generar-plantilla
```

---

## 📊 Reportes Generados

Al finalizar cada ejecución, los reportes se almacenan en la carpeta `LOGS/`:
1. **`REPORTE_NOTAS.csv`:** Columnas `Archivo, Estudiante, Modo_Evaluacion, Puntos_Obtenidos, Puntos_Totales, Nota_Final_100, Estado, Criterios_Fallados, Advertencias`. Codificado en UTF-8 con BOM para apertura directa en Excel.
2. **`LOG_NOTAS.xlsx`:** Libro formateado con colores institucionales, cabeceras profesionales, resumen general y hoja con la matriz estudiante vs criterios.
3. **`REPORTE_DETALLADO.json`:** Árbol de datos con el diagnóstico individual de cada criterio evaluado por estudiante.
4. **`REVISADOS/<nombre>_rev.xlsx`:** Cada archivo de estudiante contiene los comentarios insertados y las celdas con errores resaltadas en rojo pastel (`#FFCCCC`).

---

## 🧪 Ejecución de Pruebas Automatizadas

El proyecto incluye 16 pruebas automatizadas que validan cada módulo:
```powershell
.\.venv\Scripts\python.exe -m pytest -v tests/
```
Resultado: **16 passed in ~15s**.

---

## 📚 Documentación Adicional

- [TUTORIAL_RUBRICA.md](TUTORIAL_RUBRICA.md): Guía paso a paso para crear y personalizar la hoja `_RUBRICA` en Excel (Tablas de datos, Tablas Dinámicas en hojas nuevas, fórmulas, formatos y parámetros).

