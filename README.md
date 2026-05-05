# Muestreos SURA

Módulos VBA para sistemas de muestreo estadístico en Excel, desarrollados para auditoría interna en el contexto regulatorio SAB/SAF de SURA Investments Perú. Cada libro implementa un flujo completo de importación, cálculo de universo, selección de muestra aleatoria y exportación de resultados.

---

## Índice

- [Descripción general](#descripción-general)
- [Metodología estadística](#metodología-estadística)
- [Estructura del repositorio](#estructura-del-repositorio)
- [Módulos](#módulos)
- [Paradigma común](#paradigma-común)
- [Libros implementados](#libros-implementados)
  - [SAB — Contratos](#sab--contratos)
  - [SAB — Órdenes](#sab--órdenes)
  - [SAF — Contratos](#saf--contratos)
  - [SAF — Operaciones](#saf--operaciones)
  - [SAF — Rescates](#saf--rescates)
  - [SAF — Suscripciones](#saf--suscripciones)
- [Nombres definidos requeridos](#nombres-definidos-requeridos)
- [Requisitos](#requisitos)

---

## Descripción general

El sistema permite auditar muestras de transacciones financieras de forma reproducible y técnicamente sustentada. A partir de un archivo de datos del período (TSV o Excel), el flujo:

1. Importa y valida la estructura del archivo fuente.
2. Detecta automáticamente el período de los datos y lo escribe como etiqueta informativa.
3. Clasifica los registros en personas naturales (PN: NAT + MAN) y personas jurídicas (PJ: JUR).
4. Calcula el tamaño de muestra óptimo para cada segmento usando la fórmula de Cochran con corrección para población finita.
5. Genera números aleatorios únicos ordenados ascendentemente en una grilla de 5 columnas.
6. Exporta las filas seleccionadas a hojas formateadas listas para revisión.

El diseño es deliberadamente sin filtros de período en el cálculo: el archivo importado ya corresponde al período correcto, lo que elimina la posibilidad de errores por configuración manual de fechas.

---

## Metodología estadística

Se aplica la fórmula de Cochran con corrección para población finita:

$$n = \frac{N \cdot Z^2 \cdot p(1-p)}{(N-1) \cdot E^2 + Z^2 \cdot p(1-p)}$$

| Parámetro | Descripción | Valor típico |
|---|---|---|
| N | Tamaño total del universo (PN o PJ) | Calculado automáticamente |
| Z | Valor Z para el nivel de confianza deseado | 1.96 (95%) |
| p | Proporción esperada máxima (criterio conservador) | 0.5 |
| E | Margen de error aceptable | 0.29 (29%) |

Los parámetros Z, p y E se configuran directamente en la hoja Muestra de cada libro mediante celdas vinculadas a nombres definidos. El nivel de confianza se ingresa como porcentaje y se convierte automáticamente al valor Z correspondiente mediante fórmula.

La selección es aleatoria simple sin reemplazo: cada número del universo puede aparecer como máximo una vez en la muestra. Los números se ordenan ascendentemente antes de escribirse en la grilla para facilitar la localización de registros durante la revisión.

---

## Estructura del repositorio

```
Muestreos/
├── SAB/
│   ├── Contratos/
│   │   ├── Módulos/
│   │   │   ├── modImportarDatos.bas
│   │   │   ├── modUniversoPorTipo.bas
│   │   │   ├── modSeleccionMuestra.bas
│   │   │   ├── modExportarMuestra.bas
│   │   │   └── modEliminarDatos.bas
│   │   └── Objetos/
│   │       └── ThisWorkbook.bas
│   └── Ordenes/
│       ├── Módulos/
│       └── Objetos/
├── SAF/
│   ├── Contratos/
│   │   ├── Módulos/
│   │   └── Objetos/
│   ├── Operaciones/
│   │   ├── Módulos/
│   │   └── Objetos/
│   ├── Rescates/
│   │   ├── Módulos/
│   │   └── Objetos/
│   └── Suscripciones/
│       ├── Módulos/
│       └── Objetos/
└── README.md
```

---

## Módulos

Cada libro implementa el mismo conjunto de seis módulos con responsabilidades bien delimitadas:

### `modImportarDatos`

Punto de entrada del botón **Importar Datos**. Responsabilidades:

- Abre el archivo fuente en modo lectura.
- Detecta la hoja de datos por nombre (prioridad) y por estructura de columnas (fallback).
- Localiza el bloque de datos ignorando filas de metadatos o cabeceras decorativas encima de los headers reales.
- Evita la columna A fantasma que aparece en algunos archivos XLS cuando los datos empiezan en la columna B.
- Calcula la última fila real usando `End(xlUp)` en múltiples columnas para evitar falsos negativos por columnas dispersas.
- Valida que haya al menos una fila de datos debajo de las cabeceras; si no, muestra un error descriptivo y no importa.
- Copia los valores (no fórmulas) a la hoja destino del libro de muestreo.
- Crea el `ListObject` con nombre consistente.
- Aplica formato de fecha a las columnas de fecha detectadas.
- Llama a `AutodetectarPeriodo` para escribir la etiqueta del período.
- Llama a `TamañoPoblacion` para calcular los universos y muestras.
- En caso de archivo multi-mes: muestra advertencia con lista de meses detectados, indica si son discontinuos (posible error) y pregunta si continuar. Si el usuario cancela, limpia los datos ya importados.

### `modUniversoPorTipo`

Responsable del cálculo estadístico. Responsabilidades:

- Lee la tabla importada y clasifica cada registro como PN (NAT o MAN) o PJ (JUR) usando `NormalizarTipoPersona`.
- Cuenta el total, contPN y contPJ sin filtro de fecha (el archivo ya es del período correcto).
- Lee los parámetros Z, p y E desde los nombres definidos del libro.
- Aplica la fórmula de Cochran con corrección para población finita a cada segmento.
- Escribe los resultados en los nombres definidos `TamañoPob`, `UniversoPN`, `UniversoPJ`, `TamañoMuestraPN`, `TamañoMuestraPJ`.
- Se activa también desde `ThisWorkbook.Workbook_SheetChange` cuando se modifica la tabla importada.

### `modSeleccionMuestra`

Generador de la muestra aleatoria. Responsabilidades:

- Valida que la tabla tenga datos y que los universos sean > 0 antes de preguntar confirmación.
- Valida que los tamaños de muestra calculados sean > 0; si no, indica que se revisen los parámetros.
- Genera números aleatorios únicos sin repetición usando una `Collection` con key = número como mecanismo de deduplicación.
- Ordena los números ascendentemente (burbuja) antes de escribirlos.
- Escribe en una grilla de 5 columnas a partir de `Muestra1_PN` y `Muestra1_PJ`, copiando el formato de la celda ancla.
- Aplica bordes punteados grises a cada celda de la grilla.
- Retorna mensajes de error descriptivos por segmento si algo falla, sin mezclarlos con el mensaje de éxito.

### `modExportarMuestra`

Generador de las hojas de muestra. Responsabilidades:

- Valida que existan números en la grilla antes de proceder.
- Construye el subuniverso del tipo correspondiente (PN o PJ) recorriendo la tabla importada.
- Mapea cada número de la grilla a su fila real en el subuniverso.
- Si algún número está fuera del rango del universo actual, avisa y sugiere regenerar.
- Crea o reemplaza la hoja destino con nombre `Muestra_[Tipo]_[Segmento]_[MesAA]`.
- Copia valores directamente (sin `Copy/PasteSpecial`) para máxima velocidad.
- Agrega una columna extra con el número de posición en el universo.
- Crea un `ListObject` con estilo `TableStyleMedium7` para PN y `TableStyleMedium3` para PJ.
- Aplica formato de fecha a las columnas de fecha detectadas.
- Ajusta el ancho de columnas automáticamente.

### `modEliminarDatos`

Limpieza completa del entorno. Responsabilidades:

- Requiere doble confirmación explícita del usuario antes de ejecutar cualquier acción.
- Elimina la hoja de datos importados (no solo la limpia, la borra).
- Recolecta los nombres de todas las hojas de muestra con el prefijo correspondiente antes de borrarlas, para evitar el error clásico de modificar una colección mientras se itera.
- Resetea a 0 los nombres definidos numéricos (universos, tamaños).
- Limpia la etiqueta `PeriodoActual`.
- Limpia el contenido y los bordes de las grillas de números de muestra.
- Restaura `Application.DisplayAlerts`, `EnableEvents` y `ScreenUpdating` tanto en el flujo normal como en el manejador de errores.

### `ThisWorkbook`

Coordinación de eventos. Responsabilidades:

- `Workbook_SheetChange`: detecta cambios en la tabla importada y recalcula automáticamente el universo llamando a `TamañoPoblacion`. Monitorea solo la hoja de datos correspondiente para no dispararse con cambios en otras hojas.

---

## Paradigma común

Todos los libros comparten estos principios de diseño:

**Sin filtro de período en el cálculo:** el archivo importado ya corresponde al período deseado. `TamañoPoblacion` cuenta todos los registros sin filtrar por fecha, lo que elimina dependencias de nombres definidos de fecha (`Año`, `Mes`, `TipoInforme`) y evita la clase de error donde el universo queda en 0 por un desajuste entre los datos y los filtros.

**Autodetección de período:** `AutodetectarPeriodo` escanea la columna de fecha clave, extrae los meses presentes, los ordena cronológicamente y escribe la etiqueta en `PeriodoActual`. Para un solo mes: `"Enero 2026"`. Para un rango: `"Enero 2026 - Marzo 2026"`. Para meses discontinuos: mismo formato más advertencia de posible error.

**Normalización de tipo de persona:** `NormalizarTipoPersona` convierte `NAT`, `NATURAL`, `MAN`, `MANCOMUNADO` a `"N"` y `JUR`, `JURIDICA` a `"J"`, ignorando `Chr(160)` (espacio de no separación) que algunos archivos XLS insertan.

**Detección de columnas con `Canon`:** todas las búsquedas de nombres de columna pasan por `Canon()` que elimina tildes, `Chr(160)`, espacios, guiones y puntos antes de comparar. Esto hace que `"FECHA OPERACIÓN"`, `"FECHA OPERACION"` y `"FECHA  OPERACION"` sean equivalentes.

**Manejo de errores con paso descriptivo:** el manejador de errores captura la variable `paso` que describe exactamente qué operación estaba ejecutando cuando ocurrió el error, lo que facilita el diagnóstico sin necesidad de debugger.

---

## Libros implementados

### SAB — Contratos

Contratos de clientes del sistema SAB.

| Campo | Detalle |
|---|---|
| Formato fuente | TSV, encoding ISO-8859-1 |
| Headers | Fila 1 directa |
| Columnas | 27 columnas |
| Segmentación | PN / PJ |
| Columna tipo persona | `Tipo` |
| Columna fecha | `Fecha de Ingreso` |
| Hoja destino | `Contratos` |
| Tabla | `Contratos` |
| Ingesta | Power Query (normaliza tildes en headers, parsea fechas DDMMMYYY con años de 3 dígitos) |

---

### SAB — Órdenes

Órdenes de inversión del sistema SAB.

| Campo | Detalle |
|---|---|
| Formato fuente | TSV, encoding ISO-8859-1 |
| Headers | Fila 4 (3 filas de metadatos encima) |
| Segmentación | PN / PJ |
| Columna fecha | `Fecha` (formato DDMMMYYYY) |
| Hoja destino | `Ordenes` |
| Tabla | `Ordenes` |
| Ingesta | Power Query |

---

### SAF — Contratos

Contratos de partícipes del sistema SAF.

| Campo | Detalle |
|---|---|
| Formato fuente | Excel (.xls / .xlsx) |
| Datos desde | B11 (failsafe A1) |
| Columnas clave | `Tipo` (N/J), `Fecha de Ingreso` |
| Total columnas | 12 |
| Segmentación | PN / PJ |
| Hoja destino | `Contratos` |
| Tabla | `Contratos` |
| Ingesta | Directa (valores) |
| Preservación DNI | `NumberFormat "@"` + re-lectura `.Text` para evitar pérdida de ceros iniciales |

---

### SAF — Operaciones

Operaciones de mercado de capitales del sistema SAF.

| Campo | Detalle |
|---|---|
| Formato fuente | Excel (.xls / .xlsx) |
| Datos desde | Fila variable (metadatos encima) |
| Columnas | 23 columnas con tildes en headers |
| Segmentación | Universo único (sin PN/PJ) |
| Exclusión | Filas con `Operación = "PRECANCELACION TITULOS UNICOS"` |
| Hoja destino | `Operaciones` |
| Tabla | `Operaciones` |
| Ingesta | Power Query (`Operaciones_Raw`) — normaliza tildes, parsea fechas DD/MM/YYYY explícito para evitar inversión por locale |
| Nombres definidos | `Universo`, `TamañoMuestra`, `Muestra1` (sin segmentación PN/PJ) |

---

### SAF — Rescates

Solicitudes de rescate del sistema SAF.

| Campo | Detalle |
|---|---|
| Formato fuente | Excel (.xls) con múltiples hojas ocultas irrelevantes |
| Hoja datos | `RESCATES` (visible); fallback por estructura |
| Metadatos encima | Sí (reporte con título en filas 7-8, headers en fila 11) |
| Columnas | 29+ columnas |
| Columnas clave | `TIPOPERSONA` (NAT/MAN/JUR), `FECHA PROCESO` |
| Segmentación | PN / PJ |
| Hoja destino | `Rescates` |
| Tabla | `Rescates` |
| Ingesta | Directa (valores); `BuscarFechaOperacionEnLO` con `Canon` prioriza `FECHAPROCESO` |

---

### SAF — Suscripciones

Solicitudes de suscripción del sistema SAF.

| Campo | Detalle |
|---|---|
| Formato fuente | Excel (.xls) con múltiples hojas ocultas |
| Hoja datos | `SUBSCRIPCIONES` (nombre del sistema, con B); fallback `SUSCRIPCIONES`; fallback por estructura |
| Metadatos encima | Sí (misma estructura que Rescates, headers en fila 11) |
| Columnas | 23 columnas |
| Columnas clave | `TIPO PERSONA` (NAT/MAN/JUR), `FECHA PROCESO` |
| Segmentación | PN / PJ |
| Hoja destino | `Suscripciones` |
| Tabla | `Suscripciones` |
| Ingesta | Directa (valores) |

---

## Nombres definidos requeridos

Cada libro debe tener los siguientes nombres definidos apuntando a celdas en la hoja `Muestra`:

| Nombre | Tipo | Descripción |
|---|---|---|
| `Z` | Parámetro (editable) | Valor Z del nivel de confianza |
| `p` | Parámetro (editable) | Proporción esperada |
| `E` | Parámetro (editable) | Margen de error |
| `PeriodoActual` | Resultado (solo lectura) | Etiqueta del período detectado |
| `TamañoPob` | Resultado | Total universo |
| `UniversoPN` | Resultado | Universo personas naturales |
| `UniversoPJ` | Resultado | Universo personas jurídicas |
| `TamañoMuestraPN` | Resultado | Tamaño de muestra PN |
| `TamañoMuestraPJ` | Resultado | Tamaño de muestra PJ |
| `Muestra1_PN` | Ancla | Celda inicio grilla de números PN |
| `Muestra1_PJ` | Ancla | Celda inicio grilla de números PJ |

> Para SAF Operaciones (universo único): reemplazar `UniversoPN/PJ` y `TamañoMuestraPN/PJ` por `Universo`, `TamañoMuestra` y `Muestra1`.

---

## Requisitos

- Microsoft Excel con soporte VBA (Office 2016 o superior recomendado)
- Power Query habilitado para SAB Contratos, SAB Órdenes y SAF Operaciones
- Macros habilitadas en el libro
- Los nombres definidos deben existir antes de la primera importación; el código los escribe pero no los crea
