# 📘 Guía y Tutorial: Configuración Declarativa en la Hoja `_RUBRICA`

Este tutorial documenta cómo estructurar y configurar la hoja **`_RUBRICA`** en tus plantillas de Excel para evaluar tareas y exámenes con el **Modo B (Rúbrica Inteligente por Criterios - 0 a 100 puntos)** en `APP_REVISA_EXCEL_V2`.

---

## 1. Estructura General de la Hoja `_RUBRICA`

La hoja debe llamarse exactamente **`_RUBRICA`** (con guion bajo inicial, puede estar visible u oculta en el libro `.xlsx`).

En la **Fila 1** se colocan los siguientes 8 encabezados obligatorios:

| Columna A | Columna B | Columna C | Columna D | Columna E | Columna F | Columna G | Columna H |
| :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **ID** | **CRITERIO** | **HOJA** | **RANGO** | **TIPO_VALIDACION** | **PARAMETROS** | **PUNTOS** | **OBLIGATORIO** |

---

## 2. Ejemplo Práctico: Tabla de Datos y Tabla Dinámica en Hoja Nueva

### Escenario:
- En la hoja **`Datos`**, el estudiante tiene registros en el rango **`A1:E50`**.
- Debe convertir ese rango en una Tabla oficial de Excel y nombrarla **`TablaVentas`**.
- Debe crear una hoja nueva llamada **`Resumen_Ventas`**.
- En esa nueva hoja, debe insertar una **Tabla Dinámica** basada en `TablaVentas`, colocando el campo `Categoria` en **Filas** y la suma de `Total` en **Valores**.

### Configuración en la hoja `_RUBRICA`:

| ID | CRITERIO | HOJA | RANGO | TIPO_VALIDACION | PARAMETROS | PUNTOS | OBLIGATORIO |
| :--- | :--- | :--- | :--- | :--- | :--- | :---: | :---: |
| **CRIT_01** | Convertir rango en Tabla oficial con nombre TablaVentas | `Datos` | `A1:E50` | `tabla` | `{"nombre": "TablaVentas"}` | 40 | `SI` |
| **CRIT_02** | Insertar Tabla Dinámica en hoja nueva Resumen_Ventas | `Resumen_Ventas` | `A3` | `tabla_dinamica` | `{"fuente": "TablaVentas", "campos_fila": ["Categoria"], "campos_valor": ["Total"]}` | 60 | `SI` |

> 💡 **Nota sobre puntajes:** La suma total de la columna `PUNTOS` en este ejemplo es `40 + 60 = 100 puntos`.

---

## 3. Desglose Detallado de las Columnas

### Criterio 1: Creación y Nombre de la Tabla de Datos (`CRIT_01`)
- **`HOJA` (`Datos`):** Nombre de la hoja donde residen los datos originales.
- **`RANGO` (`A1:E50`):** El rango que abarca la tabla incluyendo sus encabezados.
- **`TIPO_VALIDACION` (`tabla`):** Instruye al motor a verificar el objeto tabla estructurada (`ListObject` de Excel).
- **`PARAMETROS` (`{"nombre": "TablaVentas"}`):** Especifica el nombre de la tabla asignado por el estudiante.
- **`PUNTOS` (`40`):** Puntaje asignado al criterio.
- **`OBLIGATORIO` (`SI`):** Si el estudiante no crea la tabla, se registrará como fallo crítico en el reporte.

---

### Criterio 2: Tabla Dinámica en Hoja Nueva (`CRIT_02`)
- **¿Cómo se exige que esté en una HOJA NUEVA?**
  - Colocando el nombre de la hoja esperada en la columna **`HOJA`** (en este ejemplo: **`Resumen_Ventas`**).
  - Si el estudiante no crea la hoja o deja un nombre genérico como `Hoja1`, el evaluador reportará de inmediato:
    > ❌ *Hoja 'Resumen_Ventas' no encontrada en el archivo del estudiante.*
- **`RANGO` (`A3`):** Celda superior izquierda donde inicia la tabla dinámica en la nueva hoja.
- **`TIPO_VALIDACION` (`tabla_dinamica`):** Activa el inspector de tablas dinámicas de Excel (XML y COM).
- **`PARAMETROS`:** Configuración en formato JSON:
  - `"fuente": "TablaVentas"` *(Verifica que se origine de la tabla estructurada).*
  - `"campos_fila": ["Categoria"]` *(Campos requeridos en Filas).*
  - `"campos_valor": ["Total"]` *(Campos requeridos en Valores).*
  - *(Opcional)* `"campos_columna": ["Mes"]` *(Campos en Columnas).*
  - *(Opcional)* `"campos_filtro": ["Año"]` *(Campos en Filtros).*
- **`PUNTOS` (`60`):** Puntaje asignado.
- **`OBLIGATORIO` (`SI`):** Requisito indispensable.

---

## 4. Parámetros Disponibles para `tabla_dinamica`

Puedes combinar cualquiera de estas claves en la columna **`PARAMETROS`**:

```json
{
  "fuente": "NombreDeLaTablaOriginal",
  "campos_fila": ["Columna1", "Columna2"],
  "campos_columna": ["Columna3"],
  "campos_valor": ["Total", "Cantidad"],
  "campos_filtro": ["Region"]
}
```

> **Reglas para escribir JSON en Excel:**
> 1. Usa siempre comillas dobles `"` para nombres de propiedades y cadenas de texto.
> 2. Si solo deseas validar que la tabla dinámica exista sin exigir campos específicos, escribe simplemente: `{}`.

---

## 5. ¿Cómo Saber Qué Poner en `TIPO_VALIDACION`? (Guía de Decisión)

Para saber qué escribir en la columna **`TIPO_VALIDACION`**, hazte esta pregunta simple: **¿Qué le estás pidiendo al estudiante que haga?**

| Si en el ejercicio pides... | En `TIPO_VALIDACION` pones: | ¿Qué puedes poner en `PARAMETROS`? |
| :--- | :--- | :--- |
| **Calcular con una fórmula o función** (ej: `=SI.CONJUNTO(...)`, `=SUMA(...)`, `=BUSCARV(...)`) | **`formula`** *(o **`funcion`**)* | `{"funciones": ["SI.CONJUNTO"]}` o `{"formula": "=..."}` |
| **Directamente el nombre de la función** | **`SI.CONJUNTO`** *(o `SUMA`, `BUSCARV`)* | `{}` *(¡El sistema lo detecta automáticamente!)* |
| **Que la celda dé un número o texto específico** (el resultado final) | **`valor`** | `{"valor": 5000, "tolerancia": 0.01}` |
| **Formato visual de moneda, porcentaje o fecha** (`$`, `₡`, `%`) | **`formato`** | `{"formato": "moneda"}` o `{"formato": "porcentaje"}` |
| **Negrita, color de celda o bordes** | **`estilo`** | `{"requiere_negrita": true, "requiere_relleno": true}` |
| **Una Tabla Dinámica** | **`tabla_dinamica`** | `{"campos_fila": ["Cat"], "campos_valor": ["Total"]}` |
| **Un Gráfico Dinámico o Hoja de Gráfico** | **`grafico_dinamico`** *(o `grafico`)* | `{"tipo_hoja": "chartsheet", "origen": "Resumen"}` |
| **Convertir datos en Tabla oficial de Excel** | **`tabla`** | `{"nombre": "TablaVentas"}` |
| **Combinar celdas** | **`celdas_combinadas`** | `{}` |
| **Una lista desplegable de validación** | **`validacion_datos`** | `{"tipo": "list"}` |

---

### Caso Especial: ¿Cómo evaluar funciones como `SI.CONJUNTO`?

Tienes **3 formas fáciles** de escribirlo en la hoja `_RUBRICA`, y el evaluador entenderá las 3 automáticamente:

#### Forma 1: Usando `formula` (Recomendada)
- **`TIPO_VALIDACION`:** `formula`
- **`PARAMETROS`:** `{"funciones": ["SI.CONJUNTO"]}`

#### Forma 2: Usando la palabra `funcion`
- **`TIPO_VALIDACION`:** `funcion`
- **`PARAMETROS`:** `{"funciones": ["SI.CONJUNTO"]}`

#### Forma 3: Escribiendo directamente el nombre de la función
- **`TIPO_VALIDACION`:** `SI.CONJUNTO`
- **`PARAMETROS`:** `{}`
*(El auditor reconoce automáticamente que `SI.CONJUNTO` es una función de Excel y la evaluará sin necesidad de configurar nada más).*

> 🌐 **Soporte Bilingüe Automático:** Si el estudiante tiene su Excel en inglés y utiliza `=IFS(...)`, el sistema lo traduce y lo califica como **100% correcto** automáticamente.

---

### Catálogo Completo de Tipos de Validación Soportados

| Tipo de Validación | Sinónimos Aceptados | Qué Verifica |
| :--- | :--- | :--- |
| `formula_flexible` | `formula`, `formulas`, `funcion`, `funciones` o nombre de función (`SUMA`, `SI.CONJUNTO`, `BUSCARV`) | Sintaxis semántica, función utilizada y referencias requeridas. |
| `valor_numerico` | `valor`, `numero`, `resultado` | Resultado final numérico o textual con tolerancia. |
| `formato_numero` | `formato`, `moneda`, `porcentaje`, `fecha` | Detección agnóstica de símbolo de moneda, porcentajes o fechas. |
| `estilo_visual_flexible` | `estilo`, `diseno`, `visual`, `negrita`, `color` | Presencia de negrita, relleno y bordes. |
| `tabla_dinamica` | `tabla_dinamica`, `pivot`, `td` | Presencia y campos de la Tabla Dinámica. |
| `grafico_dinamico` | `grafico_dinamico`, `grafico`, `chart`, `chartsheet` | Existencia del gráfico y si fue movido a hoja nueva (`ChartSheet`). |
| `tabla` | `tabla`, `tabla_oficial`, `table` | Objeto tabla estructurada y su nombre asignado. |
| `celdas_combinadas` | `celdas_combinadas`, `combinadas`, `merge` | Rango de celdas combinadas. |
| `validacion_datos` | `validacion_datos`, `lista`, `validacion` | Reglas de validación de celda o listas desplegables. |


---

### Ejemplo Específico: Gráfico Dinámico movido a Hoja Nueva (`ChartSheet`)

Cuando el ejercicio pide:
> *"Cree un gráfico dinámico a partir de la tabla dinámica de la hoja `Resumen_Ventas`, y desde el menú Análisis de gráfico dinámico > Acciones > Mover gráfico, muévalo a una **hoja nueva** llamada `Gráfico Ventas x Vendedor`."*

Se define en la hoja `_RUBRICA` de la siguiente forma:

| ID | CRITERIO | HOJA | RANGO | TIPO_VALIDACION | PARAMETROS | PUNTOS | OBLIGATORIO |
| :--- | :--- | :--- | :---: | :--- | :--- | :---: | :---: |
| **CRIT_03** | Insertar Gráfico Dinámico y mover a hoja nueva Gráfico Ventas x Vendedor | `Gráfico Ventas x Vendedor` | `-` | `grafico_dinamico` | `{"tipo_hoja": "chartsheet", "origen": "Resumen_Ventas"}` | 20 | `SI` |

**Detalles técnicos:**
- **`HOJA`:** Debe ser exactamente el nombre de la pestaña asignada: `Gráfico Ventas x Vendedor`.
- **`RANGO`:** Se coloca `-` o `A1` (en las Hojas de Gráfico no hay celdas individuales).
- **`TIPO_VALIDACION`:** `grafico_dinamico` (o `hoja_grafico`).
- **`PARAMETROS`:**
  - `"tipo_hoja": "chartsheet"`: Exige que el alumno haya usado la opción *Mover Gráfico > Hoja nueva*, y no que simplemente haya copiado el gráfico en una celda de una hoja normal.
  - `"origen": "Resumen_Ventas"`: Especifica la hoja de la Tabla Dinámica de origen.

---

### Ejemplo Específico: Operaciones Combinadas (con y sin funciones)

#### Caso 1: Operación Combinada Directa en una Celda (sin funciones)
> **Ejercicio:** *"En la celda `G3`, calcule el total con descuento aplicando la fórmula: `= D3 * E3 - D3 * E3 * 2%`"*

En la hoja `_RUBRICA` se define así:

| ID | CRITERIO | HOJA | RANGO | TIPO_VALIDACION | PARAMETROS | PUNTOS | OBLIGATORIO |
| :--- | :--- | :--- | :---: | :--- | :--- | :---: | :---: |
| **CRIT_04** | Calcular total con 2% de descuento usando operación combinada | `Ventas` | `G3` | `formula` | `{"referencias": ["D3", "E3"], "contiene": ["2%/0.02"]}` | 15 | `SI` |

**¿Por qué se configura así?**
- **`TIPO_VALIDACION`: `formula`:** Exige que el alumno escriba una fórmula que empiece con `=`, no un número estático.
- **`"referencias": ["D3", "E3"]`:** Valida que el estudiante use las celdas de Precio y Cantidad (`D3` y `E3`).
- **`"contiene": ["2%/0.02"]`:** Verifica que aplique el descuento del 2%, aceptando tanto si el alumno escribe `2%` como si escribe `0.02` (o `0,02`).

---

#### Caso 2: Operación Combinada DENTRO de un `SI` o `SI.CONJUNTO`
> **Ejercicio:** *"En la celda `F3`, si el cargo en `D3` es 'JEFE', aplique el cálculo: `=SI.CONJUNTO(D3="JEFE"; C3 * E3 - C3 * E3 * 5%; ...)`"*

En la hoja `_RUBRICA` se define así:

| ID | CRITERIO | HOJA | RANGO | TIPO_VALIDACION | PARAMETROS | PUNTOS | OBLIGATORIO |
| :--- | :--- | :--- | :---: | :--- | :--- | :---: | :---: |
| **CRIT_05** | Calcular bono condicional con SI.CONJUNTO y operación combinada | `Ventas` | `F3` | `formula` | `{"funciones": ["SI.CONJUNTO"], "referencias": ["D3", "C3", "E3"], "contiene": ["JEFE", "5%/0.05"]}` | 20 | `SI` |

**¿Por qué se configura así?**
1. **`"funciones": ["SI.CONJUNTO"]`:** Verifica que use la función requerida (acepta `SI.CONJUNTO` en español o `IFS` en inglés).
2. **`"referencias": ["D3", "C3", "E3"]`:** Comprueba que evalúe la celda condicional `D3` y calcule con `C3` y `E3`.
3. **`"contiene": ["JEFE", "5%/0.05"]`:** Valida que compare contra el texto `"JEFE"` y que aplique el porcentaje `5%` o `0.05`.

---

## 6. Reglas de Oro para la Hoja `_RUBRICA`

1. **Puntaje total exacto:** La suma de la columna `PUNTOS` debe ser `100.0`. En la interfaz web aparecerá un distintivo verde confirmando el total.
2. **Coincidencia de nombres:** Los nombres de hojas (`Resumen_Ventas`) y encabezados (`Categoria`, `Total`) deben coincidir con los datos del ejercicio.
3. **Formato de guardado:** Guarda el archivo como libro regular de Excel (`.xlsx`), sin macros.

---

## 7. Flujo de Trabajo Rápido para el Docente

1. Abre tu plantilla de ejercicio en Microsoft Excel.
2. Añade la hoja **`_RUBRICA`** con tus criterios y puntos.
3. Guarda el archivo `.xlsx`.
4. Abre el auditor con [EJECUTAR_INTERFAZ.bat](file:///c:/AAM/REPOSITORIOS/APP_REVISA_EXCEL_V2/EJECUTAR_INTERFAZ.bat).
5. Selecciona la plantilla: el sistema la pre-chequeará automáticamente y estará lista para calificar los envíos de los estudiantes con reporte detallado y marcas visuales en rojo suave (`#FFCCCC`).
