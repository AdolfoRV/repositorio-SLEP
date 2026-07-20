# Automatización de Reportes de Licencias Médicas — SLEP Los Libertadores

> **Servicio Local de Educación Pública Los Libertadores**  
> Procesador de datos históricos de licencias médicas con normalización, validación cruzada y generación de modelo estrella para Power BI.
>
> **URL de uso:** [adolforv.github.io/repositorio-SLEP/migrador.html](https://adolforv.github.io/repositorio-SLEP/migrador.html)

---

## Tabla de contenidos

- [Resumen](#resumen)
- [Demo](#demo)
- [Arquitectura](#arquitectura)
- [Stack tecnológico](#stack-tecnológico)
- [Cómo usar](#cómo-usar)
- [Entrada esperada](#entrada-esperada)
- [Salidas generadas](#salidas-generadas)
- [Dashboard de Power BI](#dashboard-de-power-bi)
- [Reglas de negocio principales](#reglas-de-negocio-principales)
- [Riesgos conocidos](#riesgos-conocidos)
- [Licencia](#licencia)

---

## Resumen

Este proyecto moderniza la gestión de licencias médicas del SLEP Los Libertadores, migrando desde una planilla Excel histórica con problemas de calidad de datos hacia un **modelo de datos normalizado** (esquema estrella) listo para análisis en Power BI.

### El problema histórico

- **Incoherencias**: datos redundantes y contradictorios entre hojas de distintos años.
- **Errores de entrada**: múltiples variantes ortográficas para un mismo establecimiento, institución o tipo de licencia ("FONASA", "fonasa", "Fonasa "…).
- **Procesos manuales**: reportar e imputar licencias exigía edición manual exhaustiva y propensa a errores.

### La solución

La herramienta funciona como un **procesador de datos inteligente** que corre 100 % en el navegador (sin servidor). El usuario solo debe cargar su **planilla madre** de licencias; el sitio descarga automáticamente el maestro de establecimientos y la plantilla de Power BI desde sus propios recursos estáticos. Al finalizar, devuelve un archivo ZIP con cinco archivos Excel normalizados y un dashboard listo para usar.

1. **Normaliza** automáticamente errores ortográficos y variantes de escritura mediante expresiones regulares y fuzzy matching.
2. **Valida cruzadamente** la información entre tablas (funcionarios, establecimientos, AFPs) y reporta anomalías.
3. **Genera reportes listos para usar**: planillas con alertas de inconsistencias, campos autorrellenados y listas desplegables para la imputación final.

---

## Demo

<video src="assets/Demo.mp4" controls width="100%" style="max-width: 800px; border-radius: 8px;"></video>

## Arquitectura

```mermaid
flowchart LR
    subgraph Navegador["Navegador del usuario (sitio Quarto)"]
        UI["Interfaz web<br/>(index.qmd / migrador.qmd<br/>+ ui.js)"]
        W["worker.js<br/>(Web Worker + Pyodide/WASM)"]
    end
    subgraph Paquete["Paquete Python 'slep'"]
        C["core.py — pipeline"]
        K["constants.py — reglas de negocio"]
        U["utils.py — normalización"]
    end
    IN1[("Planilla madre<br/>cargada por el usuario<br/>.xlsx")]
    RES[("Recursos estáticos del sitio<br/>establecimientos.xlsx + .pbit")]
    OUT[("SLEP_files.zip<br/>5 xlsx + .pbit")]
    PBI["Power BI<br/>Dashboard KPIs"]

    IN1 --> UI
    RES --> UI
    UI --> W --> C
    K --> C
    U --> C
    C --> W --> OUT
    OUT --> PBI
```

### Flujo de datos

1. El usuario abre la página del migrador y **carga únicamente su planilla madre** de licencias (`.xlsx`).
2. El sitio descarga automáticamente desde sus recursos estáticos:
   - El **maestro de establecimientos** (`assets/tables/establecimientos.xlsx`).
   - La **plantilla de Power BI** (`assets/Dashboard_Licencias.pbit`), si está disponible.
3. El *worker* invoca `slep.procesar(...)` y muestra en pantalla el **log** de clasificación e inconsistencias.
4. El usuario descarga `SLEP_files.zip` con los cinco archivos normalizados (+ dashboard).
5. Los archivos se usan como nueva planilla madre de imputación y como fuente del dashboard de Power BI.

### Decisiones de diseño clave

- **Sin servidor**: al correr en Pyodide/WASM, el procesamiento es 100 % local. Los archivos del usuario **no salen de su computador**.
- **Excel como interfaz de datos**: en vez de imponer un sistema nuevo, la herramienta lee y escribe el formato que el equipo ya domina.
- **Reglas declarativas separadas del código**: todas las expresiones regulares y catálogos viven en `constants.py`, de modo que ajustar una regla de negocio no requiere tocar la lógica.

---

## Stack tecnológico

| Capa | Tecnología | Rol |
|---|---|---|
| **Interfaz y documentación** | [Quarto](https://quarto.org/) (HTML/CSS/JS) | Sitio web del proyecto: página de inicio, documentación y pantalla del migrador (`migrador.qmd`) |
| **Ejecución en el navegador** | JavaScript (`assets/ui.js`, `assets/worker.js`) + Pyodide | El Python corre **dentro del navegador** compilado a WebAssembly |
| **Motor de procesamiento** | Python 3, paquete `slep` (`scripts/slep/`) | Toda la lógica de migración, normalización y generación de archivos |
| **Almacenamiento** | Excel (`.xlsx`) | Única fuente de entrada y formato de salida |
| **Análisis** | Power BI (`.pbit`) | Dashboard de KPIs de licencias que consume los cinco archivos generados |

---

## Cómo usar

### Requisitos

- Navegador moderno con soporte para Web Workers y WebAssembly (Chrome, Edge, Firefox, Safari).
- Conexión a Internet para descargar Pyodide (~20-30 segundos la primera vez).

### Pasos

1. Abrir la página del migrador:  
   **[adolforv.github.io/repositorio-SLEP/migrador.html](https://adolforv.github.io/repositorio-SLEP/migrador.html)**
2. Arrastrar o seleccionar la **planilla madre** de licencias (`.xlsx`).  
   > No es necesario cargar nada más: el maestro de establecimientos y la plantilla Power BI se descargan automáticamente desde el sitio.
3. Presionar **"Procesar y descargar"**.
4. Revisar el log en pantalla: muestra folios repetidos, valores clasificados, inconsistencias detectadas y establecimientos nuevos.
5. Descargar `SLEP_files.zip` con los cinco archivos normalizados.

> **Nota**: el procesamiento tiene un timeout de 2 minutos. Si el worker no responde, se ofrece la opción de reintentar.

---

## Entrada esperada

### Planilla madre de licencias (`.xlsx`)

Único archivo que el usuario debe cargar. Debe contener las siguientes hojas:

| Hoja | Obligatoria | Estructura esperada |
|---|---|---|
| `DATOS` | Sí | Encabezados en **fila 1**: `RUN`, `Nombre`, `Fecha Nacimiento`, `Sexo`, `Estado Civil`, `Dirección`, `Comuna`, `Teléfono`, `Teléfono Emergencia`, `Nacionalidad`, `Formación Profesional`, `Cargo`, `Centro de Costo`. Un funcionario por fila. |
| `LM01-2024` | Sí | Encabezados en **fila 2**; se usa su columna `Unidad` (desde fila 3) como fuente adicional de establecimientos. |
| `LM*` (resto) | Al menos una | Hojas de hechos. La fila de encabezado se detecta automáticamente entre las filas 1 y 2 buscando una celda que contenga "rut". |

**Tolerancias del lector** (nombres de columna alternativos aceptados en las hojas de hechos):

| Concepto | Encabezados aceptados |
|---|---|
| Folio | `Folio licencia`, `Folio Minsal` |
| Fechas | `Fecha Inicio` / `Fech. Inicio`, `Fecha Termino` / `Fech. Termino` |
| Días | `Días LM`, `Días Lic` |
| Institución | `Institución Salud`, `Institucion Salud` |
| Estado | `Resolución Médica`, `Resolucion Medica` (con *fallback* a `Estado`) |
| Establecimiento | `Estableciemiento` *(sic, error histórico de la planilla)*, `Establecimiento`, `Unidad`, `Centro de Costo`, `Lugar`, `Sede`, `Ubicacion` |
| AFP | `A.F.P.` |

> **Nota técnica**: los libros se leen con `data_only=True`, es decir, se toma el **último valor calculado** por Excel de cada celda, no las fórmulas.

### Recursos descargados automáticamente por el sitio

El usuario **no necesita cargar** los siguientes archivos; el migrador los descarga automáticamente desde los recursos estáticos del sitio:

| Recurso | Ubicación en el sitio | Descripción |
|---|---|---|
| Maestro de establecimientos | `assets/tables/establecimientos.xlsx` | Primera hoja del libro; la fila de encabezado se detecta dentro de las primeras 4 filas buscando la celda `Tipo`. Columnas esperadas: `Tipo`, `Nombre establecimiento`, `Comuna`, `Dirección`, `Telefono`, `Sitio web`. Nombres duplicados se ignoran (gana la primera aparición). |
| Plantilla Power BI | `assets/Dashboard_Licencias.pbit` | Dashboard de KPIs; se incluye en el ZIP de salida si está disponible. |

---

## Salidas generadas

`procesar()` devuelve un paquete de archivos (todo se entrega comprimido en **`SLEP_files.zip`**):

| Archivo | Tipo | Contenido |
|---|---|---|
| `01_Dim_Funcionario.xlsx` | Dimensión | Funcionarios (incluye placeholders para RUT no encontrados), ordenados por nombre. |
| `02_Dim_Establecimiento.xlsx` | Dimensión | Maestro de establecimientos + los **nuevos** detectados (Tipo "Otro"). |
| `03_Dim_AFP.xlsx` | Dimensión | Combinaciones AFP + Tasa observadas (con centinela `-1` para tasas desconocidas). |
| `04_Hechos_Licencias.xlsx` | **Hechos** | Licencias migradas deduplicadas **+ 40 filas en blanco** para imputar nuevas, con autorrelleno y validaciones. |
| `05_Hechos_Descuentos.xlsx` | Hechos | Descuentos en formato largo (Folio, RUT, Período `YYYY-MM`, Monto, Fuente). |
| `Dashboard_Licencias.pbit` | Plantilla | Dashboard de Power BI (solo si la plantilla está disponible en el sitio). |

Todos los `.xlsx` salen con formato corporativo: encabezado azul congelado, bordes, autoancho y **tabla nativa de Excel** (filtros incluidos), listos para conectar a Power BI.

### Modelo estrella

```mermaid
flowchart TD
    F["Dim_Funcionario<br/>(RUT)"]
    E["Dim_Establecimiento<br/>(Establecimiento)"]
    A["Dim_AFP<br/>(AFP)"]
    H["Hechos_Licencias<br/>(una fila por licencia)"]
    D["Hechos_Descuentos<br/>(una fila por descuento mensual)"]
    F -->|RUT| H
    E -->|Establecimiento| H
    A -->|AFP| H
    H -->|Folio Licencia| D
```

Relaciones en Power BI:

- `Hechos_Licencias.RUT → Dim_Funcionario.RUT`
- `Hechos_Licencias.Establecimiento → Dim_Establecimiento.Establecimiento`
- `Hechos_Licencias.A.F.P. → Dim_AFP.AFP`
- `Hechos_Descuentos.Folio Licencia → Hechos_Licencias.Folio Licencia`

---

## Dashboard de Power BI

El archivo `Dashboard_Licencias.pbit` es una **plantilla de Power BI** que convierte los cinco archivos Excel generados en un dashboard interactivo de 6 páginas con KPIs, alertas y análisis temporal. No contiene datos propios: al abrirla, solicita la ruta de la carpeta donde se descomprimió `SLEP_files.zip` y lee los archivos en ese momento.

### Requisitos

- **Power BI Desktop** (versión reciente, descargable gratuitamente desde [powerbi.microsoft.com](https://powerbi.microsoft.com/)).
- Los cinco archivos Excel (`01_` a `05_`) en la misma carpeta que el `.pbit`.

### Cómo abrir el dashboard

1. Descomprimir `SLEP_files.zip` en una carpeta.
2. Abrir `Dashboard_Licencias.pbit` con Power BI Desktop.
3. Cuando aparezca el cuadro de diálogo, ingresar la **ruta de la carpeta** donde están los cinco Excel (debe terminar en `\` o `/`, según el sistema operativo).
4. Power BI cargará los datos y mostrará el informe listo para usar.

### Las 6 páginas del informe

| Página | Qué muestra | Uso típico |
|---|---|---|
| **Resumen** | KPIs globales: total de licencias, autorizadas, rechazadas/reducidas, días solicitados, total recuperado. Gráficos de evolución mensual, distribución por tipo de licencia e institución de salud. | Visión ejecutiva del estado general. |
| **Establecimientos** | Comparación entre escuelas, liceos, jardines y unidades: licencias por establecimiento, funcionarios con licencia, promedio de días. | Identificar establecimientos con mayor incidencia de licencias. |
| **Funcionarios** | Tabla de todos los funcionarios con licencia, con filtros por nombre y estado de resolución. Muestra días aprobados en ventana de 24 meses y estado de alerta. | Búsqueda puntual de una persona y detección temprana de casos críticos. |
| **Ficha Funcionario** | Página de **drillthrough**: se abre automáticamente al hacer clic derecho → "Obtener detalles" sobre un funcionario en las páginas "Funcionarios" o "Alertas". Muestra ficha personal, evolución mensual de días y listado detallado de cada licencia. | Análisis individual profundo. |
| **Alertas** | Funcionarios que superan **180 días aprobados en una ventana móvil de 24 meses**. Tarjetas con cantidad de funcionarios en alerta, porcentaje y días acumulados. Tabla filtrada solo a los casos que requieren gestión. | Gestión proactiva de ausentismo prolongado. |
| **Recuperación** | Montos recuperados vs. esperados por el sistema, brecha Sistema vs. Pagado, descuentos mensuales aplicados en remuneraciones. | Control financiero de subsidios y descuentos. |

### Filtros globales

- **Segmentador de establecimiento**: aparece en las 6 páginas, sincronizado. Permite filtrar todo el informe por uno o varios establecimientos.
- **Navegador de páginas**: menú de pestañas para moverse entre las 6 páginas.

### Medidas clave

| Medida | Descripción |
|---|---|
| `Total Licencias` | Cantidad de licencias en el contexto de filtro actual. |
| `Licencias Autorizadas` | Incluye estados "Autorizada" y "Ampliada". |
| `Dias Solicitados` | Suma de días de licencia (columna `Dias LM`). |
| `Total Recuperado` | Suma de `Total Pagado` (lo efectivamente reembolsado). |
| `Brecha Sistema vs Pagado` | Diferencia entre `Total Sistema` y `Total Pagado`. |
| `Funcionarios en Alerta` | Funcionarios con más de 180 días aprobados en los últimos 24 meses. |
| `Total Descuentos` | Suma de descuentos mensuales aplicados en remuneraciones. |

> **Nota sobre alertas**: la ventana de 24 meses es teórica; como la base tiene cobertura densa recién desde enero de 2025, la ventana "efectiva" con datos reales es algo menor mientras se acumula más historial.

### Tip de uso

Si se agregan nuevas licencias en las 40 filas en blanco de `04_Hechos_Licencias.xlsx`, solo hay que guardar el archivo, volver a Power BI y presionar **"Actualizar"** (o `Ctrl + Shift + E`) para que el dashboard refleje los cambios sin necesidad de regenerar todo desde la web.

---

## Reglas de negocio principales

Cada regla tiene un identificador **RB-*** que aparece también como comentario en el código Python (`# RB-04`, etc.), trazable desde el documento técnico.

| ID | Regla | Descripción |
|---|---|---|
| RB-01 | Normalización canónica de texto | `utils.norm()` — quita tildes, minúsculas, espacios múltiples, normaliza "n°". |
| RB-02 | Normalización de RUT | `utils.norm_rut()` — reconstruye como `NNNNNNNN-DV` en mayúsculas; **no valida DV** para conservar trazabilidad. |
| RB-03 | Errores de Excel = vacío | `#N/A`, `#REF!`, `#VALUE!`, etc. se tratan como vacío. |
| RB-04 | Clasificación en cascada | Regex → fuzzy (≥ 0,6) → "REVISAR: no reconocido". |
| RB-05 | Resoluciones legacy | Números de resolución en columna "Resolución Médica" se dejan vacíos con marca informativa. |
| RB-06 | Parseo AFP "Nombre (tasa)" | Separa nombre y tasa de cotización entre paréntesis. |
| RB-07 | Imputación de tasas AFP | Hereda tasa conocida del histórico; si no existe, centinela `-1`. |
| RB-08 | Establecimientos: 4 niveles | Exacto → regex → fuzzy (≥ 0,82) → nuevo (Tipo "Otro"). |
| RB-09 | Identificación del funcionario | RUT exacto → nombre difuso (≥ 0,85) → placeholder. |
| RB-10 | Deduplicación de hechos | Clave `(RUT, Folio, Fecha Inicio)`; gana la fuente de **mayor año**. |
| RB-11 | Montos dobles | Sistema vs. Pagado (2024 / 2025-2026) en 8 columnas. |
| RB-12 | Descuentos por período | Formato ancho (`MONTO DESCONTADO MES AÑO`) → largo (`YYYY-MM`). |
| RB-13 | Fecha término < inicio | Se conserva la fila pero se anota inconsistencia. |
| RB-14 | Detección automática de encabezados | Hojas de hechos: filas 1-2 buscando "rut"; maestro: primeras 4 filas buscando "Tipo". |
| RB-15 | Filas sin RUT ni folio | Se omiten (no trazables). |
| RB-16 | Lectura con `data_only=True` | Se leen valores calculados, no fórmulas. |

---

## Riesgos conocidos

- **Correcciones fuzzy no son verdad absoluta**: todo lo marcado "Corregido (revisar)" debe auditarse; un umbral de 0,6 puede atraer valores a la categoría equivocada en textos muy cortos.
- **El orden de las regex es significativo**: un patrón amplio declarado antes puede "capturar" un texto que debía ir a otra categoría.
- **Duplicados por folio compartido**: si dos licencias *distintas* compartieran RUT + Folio + Fecha Inicio, la deduplicación (RB-10) conservaría solo una. El log de folios repetidos es la herramienta para detectar estos casos.
- **Dependencia de convenciones de la planilla madre**: nombres de hoja `LM*`, fila de encabezado con "rut", columnas con los alias conocidos. Una planilla que se salga de estas convenciones puede requerir ajustes.
- **Tasa centinela `-1`**: indica AFP reconocida cuya tasa nunca apareció en el histórico; debe completarse en `Dim_AFP` antes de usar montos calculados con ella.
- **El RUT no valida dígito verificador** (RB-02): un RUT mal digitado en el origen se propaga tal cual para no perder trazabilidad.
- **Umbral de alerta fijo en el dashboard**: los 180 días y la ventana de 24 meses están escritos como literales en las fórmulas DAX; si la política de alerta cambia, el dashboard debe editarse manualmente en Power BI Desktop.
- **Cobertura de datos desigual en alertas**: como la base tiene datos densos recién desde enero de 2025, las alertas de "24 meses" en la práctica cubren un período menor mientras no se acumule más historial.

---

## Licencia

Proyecto interno del Servicio Local de Educación Pública Los Libertadores. Uso exclusivo para la gestión de licencias médicas de la organización.

---

*Ante cualquier discrepancia entre este README y el código fuente, el código es la fuente de verdad.*