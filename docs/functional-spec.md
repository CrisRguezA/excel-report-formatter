# Especificación Funcional — excel-report-formatter

## 1. Resumen del proyecto

**excel-report-formatter** es un módulo de automatización que transforma un dataset de ventas ya limpio en un reporte Excel profesional, listo para ser usado por equipos de ventas en su operativa semanal.

**Problema que resuelve:**
Los equipos de ventas reciben datos limpios pero sin estructura visual. Formatear manualmente un informe semanal — anchos de columna, colores por estado, totales, formatos de fecha y moneda — consume tiempo y genera inconsistencias entre semanas o entre personas.

**Objetivo principal:**
Automatizar la capa de presentación del dato limpio, produciendo un entregable reproducible y consistente sin intervención manual.

**Lugar en el pipeline:**
Este módulo es el tercer eslabón de un pipeline de automatización Excel:
```
excel-sales-consolidator → excel-data-cleaner → excel-report-formatter
```
## 1.1 Valor para negocio

Este módulo permite:

- Reducir el tiempo manual de generación de reportes semanales
- Garantizar consistencia visual entre informes
- Evitar errores humanos en formatos y cálculos
- Entregar reportes listos para toma de decisiones sin intervención adicional

Impacto esperado:
- Ahorro de tiempo operativo
- Mejora en la calidad de reporting
- Mayor confianza en los datos presentados


---

## 2. Usuarios y roles

| Usuario | Descripción | Interacción con el sistema |
|---|---|---|
| **Responsable de ventas** | Recibe el reporte semanal | Abre el Excel generado, lo usa para seguimiento operativo |
| **Técnico / Data Engineer** | Ejecuta el pipeline | Lanza el script, revisa que el output sea correcto |
| **Dirección / Stakeholder** | Toma decisiones basadas en el reporte | Consulta el Excel para análisis de negocio |

> En la fase actual (local), el técnico ejecuta el script manualmente. En la fase cloud futura, la ejecución será automática.

---

## 3. Casos de uso

| ID | Actor | Acción | Resultado esperado |
|---|---|---|---|
| CU-01 | Técnico | Ejecutar el formatter con un Excel limpio | Se genera `weekly_sales_report.xlsx` en `outputs/` |
| CU-02 | Técnico | Ejecutar con ruta de input personalizada | El sistema acepta la ruta vía `--input` y procesa el archivo indicado |
| CU-03 | Técnico | Ejecutar con columnas incorrectas o faltantes | El sistema lanza un error descriptivo y detiene la ejecución |
| CU-04 | Responsable de ventas | Abrir el reporte generado | Ve los datos formateados con colores por estado, totales y filtros activos |
| CU-05 | Responsable de ventas | Filtrar por comercial o producto | Usa los filtros nativos de Excel sobre la tabla generada |

---

## 4. Flujos de uso

### Flujo principal: generación del reporte semanal

```
1. Técnico recibe el Excel limpio (output de excel-data-cleaner)
2. Coloca el archivo en data/raw/
3. Ejecuta: python src/main.py
4. El sistema valida que el archivo tiene las columnas requeridas
5. Se genera weekly_sales_report.xlsx en outputs/
6. Técnico entrega el archivo al responsable de ventas
7. Responsable de ventas utiliza el reporte para seguimiento y toma de decisiones
```

### Flujo alternativo: error por columnas faltantes

```
1. Técnico ejecuta el script con un archivo incorrecto
2. El sistema detecta columnas faltantes
3. Muestra mensaje de error con las columnas que faltan
4. Detiene la ejecución sin generar output
```

---

## 5. Requisitos funcionales

### Hoja Weekly_Report

| ID | Requisito |
|---|---|
| RF-01 | El reporte debe incluir todas las filas del dataset limpio de entrada |
| RF-02 | La primera fila debe mostrar el título del reporte y la fecha de generación |
| RF-03 | La cabecera de columnas debe estar visualmente diferenciada (fondo azul oscuro, texto blanco) |
| RF-04 | Las filas deben colorearse según el valor de `estado`: Cerrado → verde, Pendiente → amarillo, Cancelado → rojo |
| RF-05 | La celda `certificacion` debe tener color propio independiente del color de fila |
| RF-06 | La columna `fecha_venta` debe mostrarse en formato DD/MM/YYYY |
| RF-07 | Las columnas `importe` y `precio_m3` deben mostrarse con formato moneda (€) |
| RF-08 | Debe existir una fila de totales al final con: importe total, número de ventas y % de ventas cerradas |
| RF-09 | Los filtros de Excel deben estar activos desde la fila de cabecera |
| RF-10 | La primera fila visible al hacer scroll debe ser siempre la cabecera (freeze panes) |
| RF-11 | El sistema debe ajustar automáticamente el ancho de columnas al contenido |

### Hoja Report_Info

| ID | Requisito |
|---|---|
| RF-12 | Debe existir una segunda hoja con metadatos del reporte |
| RF-13 | Los metadatos deben incluir: nombre del archivo fuente, fecha de generación, total de filas, ventas cerradas, ventas certificadas |

### Validación de entrada

| ID | Requisito |
|---|---|
| RF-14 | El sistema debe verificar que el archivo de entrada existe antes de procesarlo |
| RF-15 | El sistema debe verificar que el archivo contiene las 12 columnas estándar del pipeline |
| RF-16 | Si faltan columnas, el sistema debe indicar exactamente cuáles faltan y detener la ejecución |

---

## 6. Reglas de negocio

| Regla | Descripción |
|---|---|
| RN-01 | El color de fila tiene prioridad sobre cualquier otro color: lo determina `estado` |
| RN-02 | El color de la celda `certificacion` es independiente del color de fila |
| RN-03 | El formatter no modifica el contenido de los datos, únicamente su representación visual |
| RN-04 | La validación de calidad del dato es responsabilidad del módulo `excel-data-cleaner` |
| RN-05 | Si no hay valor de `estado` reconocido, se aplica zebra striping como fallback |

---

## 7. Consideraciones técnicas

- **Input requerido:** archivo `.xlsx` con 12 columnas estándar (output de `excel-data-cleaner`)
- **Output generado:** archivo `.xlsx` con dos hojas: `Weekly_Report` y `Report_Info`
- **Motor de escritura:** `xlsxwriter` (creación desde cero con formato completo)
- **Motor de lectura:** `pandas` + `openpyxl` (requerido internamente por `pandas.read_excel()`)
- **Ejecución:** script Python desde terminal con CLI (`--input`, `--output`)
- **Sin dependencias de red:** el módulo opera completamente en local

---

## 8. Criterios de aceptación

| ID | Criterio |
|---|---|
| CA-01 | Dado un Excel limpio válido, el sistema genera `weekly_sales_report.xlsx` sin errores |
| CA-02 | El archivo generado contiene exactamente dos hojas: `Weekly_Report` y `Report_Info` |
| CA-03 | Todas las filas con `estado = Cerrado` tienen fondo verde en la hoja Weekly_Report |
| CA-04 | La fila de totales muestra el importe total, número de ventas y % de cerradas correctamente |
| CA-05 | La columna `fecha_venta` se visualiza como fecha (DD/MM/YYYY) y no como texto |
| CA-06 | Dado un archivo con columnas faltantes, el sistema muestra un error descriptivo y no genera output |
| CA-07 | El script es ejecutable desde terminal con `python src/main.py` sin argumentos adicionales |

---

## 9. Futuras mejoras

- Integración en pipeline automatizado con ingestión desde AWS S3
- Generación de reportes analíticos (KPIs, gráficos, tablas pivot) sobre datos en RDS
- Soporte para plantillas Excel personalizadas por cliente
- Parametrización de reglas visuales vía fichero de configuración
- Adaptación del reporte a distintos perfiles de usuario (ventas, dirección, operaciones)

---

## 10. Supuestos

- El archivo de entrada ha sido previamente validado y limpiado
- Las columnas siguen el esquema estándar definido en el pipeline
- No se contemplan datos corruptos o inconsistentes en esta fase