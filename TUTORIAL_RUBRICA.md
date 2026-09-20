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

## 5. Catálogo de Tipos de Validación Disponibles en el Modo B

| Tipo de Validación | Descripción | Ejemplo de Parámetros |
| :--- | :--- | :--- |
| `formula` / `formula_flexible` | Valida fórmulas semánticas bilingües (ES/EN) y referencias requeridas | `{"funciones": ["SUMA", "SUM"], "referencias": ["C2:C20"]}` |
| `valor` / `valor_numerico` | Valida resultados calculados con margen de tolerancia | `{"valor": 1500.50, "tolerancia": 0.01}` |
| `formato` / `formato_numero` | Valida formato de Moneda, Porcentaje, Fecha, etc. | `{"formato": "moneda"}` o `{"formato": "porcentaje"}` |
| `estilo` / `estilo_visual_flexible` | Valida negrita, bordes o color de fondo | `{"requiere_negrita": true, "requiere_bordes": true}` |
| `tabla_dinamica` | Valida existencia y campos de tablas dinámicas | `{"campos_fila": ["Categoria"], "campos_valor": ["Total"]}` |
| `grafico_dinamico` | Valida Gráfico Dinámico y Hoja de Gráfico (`ChartSheet`) movida | `{"tipo_hoja": "chartsheet", "origen": "Resumen_Ventas"}` |
| `tabla` | Valida objeto Tabla oficial y su nombre | `{"nombre": "TablaVentas"}` |
| `celdas_combinadas` | Valida que el rango esté combinado | `{}` |
| `validacion_datos` | Valida listas desplegables o reglas de celda | `{"tipo": "list"}` |

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
