"""
formatter.py

ES:
Genera un reporte semanal Excel profesional a partir de un dataset ya limpio.
Produce una hoja principal (Weekly_Report) con título, tabla formateada,
colores condicionales por estado y certificación, fila de totales,
y una hoja secundaria (Report_Info) con metadatos.

Implementado con pandas + xlsxwriter.
Arquitectura abierta a extensiones con openpyxl si se requiere post-edición.

Contrato del pipeline:
    run_formatter(input_path: str, output_path: str) -> None
    input_path  : ruta al Excel limpio (output del excel-data-cleaner)
    output_path : ruta donde se guarda el reporte semanal

EN:
Generates a professional weekly Excel report from an already-clean dataset.
Produces a main sheet (Weekly_Report) with title, formatted table,
conditional colours by estado and certificacion, totals row,
and a secondary sheet (Report_Info) with report metadata.

Implemented with pandas + xlsxwriter.
Open to openpyxl extensions for post-processing if needed.
"""

# ------------------------------------------------------------------
# IMPORTS
# ------------------------------------------------------------------

import pandas as pd
import xlsxwriter
from datetime import datetime
from pathlib import Path


# ------------------------------------------------------------------
# STYLE CONSTANTS
# ------------------------------------------------------------------

HEADER_BG        = "#1F4E79"
HEADER_FONT_COL  = "#FFFFFF"
ALT_BG           = "#EBF3FB"
TITLE_BG         = "#D6E4F0"
TOTALS_BG        = "#1F4E79"
TOTALS_FONT_COL  = "#FFFFFF"

# Colores condicionales por estado
# Regla de negocio: estado tiene prioridad sobre certificacion sobre zebra
# Documentar en README si se añaden nuevos valores al pipeline
COLOR_CERRADO    = "#C6EFCE"   # verde suave
COLOR_PENDIENTE  = "#FFEB9C"   # amarillo suave
COLOR_CANCELADO  = "#FFC7CE"   # rojo suave — contemplado aunque no existe en v1 del cleaner

# Colores condicionales por certificacion (solo si estado no aplica color)
COLOR_FSC        = "#E2EFDA"   # verde muy suave
COLOR_PEFC       = "#EBF3FB"   # azul muy suave
COLOR_CE         = "#FFF2CC"   # amarillo muy suave
COLOR_SIN_CERT   = "#F2F2F2"   # gris suave

# Columnas requeridas — contrato con excel-data-cleaner
# NOTE (Gepeta): separar en obligatorias/opcionales si el formatter
# se reutiliza en otros pipelines con esquemas distintos
REQUIRED_COLUMNS = [
    "id_venta", "cliente", "fecha_venta", "producto", "tipo_madera",
    "certificacion", "cantidad_m3", "precio_m3", "importe", "estado",
    "comercial", "pais"
]


# ------------------------------------------------------------------
# LOAD & VALIDATE
# ------------------------------------------------------------------

def load_clean_excel(input_path: str) -> pd.DataFrame:
    """Carga el Excel limpio y devuelve un DataFrame."""
    path = Path(input_path)
    if not path.exists():
        raise FileNotFoundError(f"Archivo no encontrado: {path}")
    df = pd.read_excel(path)
    print(f"  ✓ Cargado: {path.name} — {df.shape[0]} filas × {df.shape[1]} columnas")
    return df


def validate_required_columns(df: pd.DataFrame) -> None:
    """Verifica que el DataFrame tenga las columnas esperadas del pipeline."""
    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError(f"Columnas faltantes en el input: {missing}")
    print(f"  ✓ Columnas validadas")


# ------------------------------------------------------------------
# FORMAT HELPER
# ------------------------------------------------------------------

def _fmt(workbook, bg: str, extra: dict | None = None) -> object:
    """
    Crea un formato xlsxwriter con base común + overrides.
    Usa None como valor por defecto para evitar el riesgo del dict mutable.
    """
    extra = extra or {}
    base = {
        "font_name": "Arial", "font_size": 10,
        "border": 1, "valign": "vcenter", "bg_color": bg
    }
    return workbook.add_format({**base, **extra})


# ------------------------------------------------------------------
# FORMAT CACHE
# ------------------------------------------------------------------

def build_format_cache(workbook) -> dict:
    """
    Crea todos los formatos de celda una sola vez y los devuelve en un dict.
    Evita llamar a workbook.add_format() dentro del bucle fila/columna,
    lo que escalaría mal en archivos grandes.

    Clave del dict: (bg_color, num_format, align)
    """
    # Todos los fondos posibles
    backgrounds = [
        COLOR_CERRADO, COLOR_PENDIENTE, COLOR_CANCELADO,
        COLOR_FSC, COLOR_PEFC, COLOR_CE, COLOR_SIN_CERT,
        ALT_BG, "#FFFFFF"
    ]

    # Combinaciones de (num_format, align) por tipo de columna
    type_variants = [
        (None,           "left"),    # texto izquierda
        (None,           "center"),  # texto centro
        ("DD/MM/YYYY",   "center"),  # fecha
        ("#,##0.00 €",   "center"),  # euro
        ("#,##0.00",     "center"),  # decimal
        ("0",            "center"),  # entero
    ]

    cache = {}
    for bg in backgrounds:
        for (num_format, align) in type_variants:
            extra = {"align": align}
            if num_format:
                extra["num_format"] = num_format
            key = (bg, num_format, align)
            cache[key] = _fmt(workbook, bg, extra)

    return cache


# ------------------------------------------------------------------
# ROW COLOR LOGIC
# ------------------------------------------------------------------

def _row_bg(estado: str, cert: str, row_idx: int) -> str:
    """
    Determina el color de fondo de una fila según esta prioridad:
      1. estado  (Cerrado / Pendiente / Cancelado)
      2. certificacion (FSC / PEFC / CE / Sin certificación)
      3. zebra striping (alterno)

    Regla de negocio explícita — documentar en README si cambia.
    """
    estado_color = {
        "Cerrado":   COLOR_CERRADO,
        "Pendiente": COLOR_PENDIENTE,
        "Cancelado": COLOR_CANCELADO,
    }
    cert_color = {
        "FSC":               COLOR_FSC,
        "PEFC":              COLOR_PEFC,
        "CE":                COLOR_CE,
        "Sin certificación": COLOR_SIN_CERT,
    }

    if estado in estado_color:
        return estado_color[estado]
    if cert in cert_color:
        return cert_color[cert]
    return ALT_BG if row_idx % 2 == 0 else "#FFFFFF"


# ------------------------------------------------------------------
# TOTALS SUMMARY
# ------------------------------------------------------------------

def build_totals_summary(df: pd.DataFrame) -> str:
    """
    Construye el texto resumen de la fila de totales.
    Separado de add_totals_row() para desacoplar cálculo y presentación.
    """
    total_ventas   = len(df)
    total_cerradas = int((df["estado"] == "Cerrado").sum())
    pct_cerradas   = total_cerradas / total_ventas * 100 if total_ventas > 0 else 0
    return f"{total_ventas} ventas ({total_cerradas} cerradas — {pct_cerradas:.0f}%)"


# ------------------------------------------------------------------
# WRITE TITLE ROW
# ------------------------------------------------------------------

def _write_title(ws, workbook, n_cols: int) -> None:
    """Escribe la fila de título con fecha de generación."""
    title_fmt = workbook.add_format({
        "font_name": "Arial", "font_size": 13, "bold": True,
        "bg_color": TITLE_BG, "font_color": "#1F4E79",
        "align": "left", "valign": "vcenter"
    })
    fecha_gen = datetime.now().strftime("%d/%m/%Y")
    ws.merge_range(0, 0, 0, n_cols - 1,
                   f"Reporte Semanal de Ventas — Generado: {fecha_gen}", title_fmt)
    ws.set_row(0, 24)


# ------------------------------------------------------------------
# WRITE HEADER ROW
# ------------------------------------------------------------------

def _write_header(ws, workbook, df: pd.DataFrame) -> None:
    """Escribe la fila de cabecera con estilo."""
    header_fmt = workbook.add_format({
        "font_name": "Arial", "font_size": 11, "bold": True, "border": 1,
        "bg_color": HEADER_BG, "font_color": HEADER_FONT_COL,
        "align": "center", "valign": "vcenter"
    })
    for col_idx, col_name in enumerate(df.columns):
        ws.write(1, col_idx, col_name, header_fmt)
    ws.set_row(1, 22)


# ------------------------------------------------------------------
# WRITE DATA ROWS
# ------------------------------------------------------------------

def apply_column_formats(ws, df: pd.DataFrame, fmt_cache: dict) -> None:
    """
    Escribe las filas de datos usando el caché de formatos.
    El color de fondo se determina por _row_bg() — estado > cert > zebra.
    """
    col_type = {
        "id_venta":      ("0",          "center"),
        "cliente":       (None,         "left"),
        "fecha_venta":   ("DD/MM/YYYY", "center"),
        "producto":      (None,         "left"),
        "tipo_madera":   (None,         "left"),
        "certificacion": (None,         "left"),
        "cantidad_m3":   ("#,##0.00",   "center"),
        "precio_m3":     ("#,##0.00 €", "center"),
        "importe":       ("#,##0.00 €", "center"),
        "estado":        (None,         "center"),
        "comercial":     (None,         "left"),
        "pais":          (None,         "left"),
    }

    col_names = list(df.columns)

    for row_idx, (_, row) in enumerate(df.iterrows()):
        estado  = str(row.get("estado", ""))
        cert    = str(row.get("certificacion", ""))
        bg      = _row_bg(estado, cert, row_idx)
        excel_row = row_idx + 2

        for col_idx, col_name in enumerate(col_names):
            value = row[col_name]
            try:
                if pd.isna(value):
                    value = None
            except (TypeError, ValueError):
                pass

            num_fmt, align = col_type.get(col_name, (None, "left"))
            cell_fmt = fmt_cache[(bg, num_fmt, align)]

            if value is None:
                ws.write_blank(excel_row, col_idx, None, cell_fmt)
            else:
                ws.write(excel_row, col_idx, value, cell_fmt)


# ------------------------------------------------------------------
# WRITE TOTALS ROW
# ------------------------------------------------------------------

def add_totals_row(ws, workbook, df: pd.DataFrame) -> None:
    """Escribe la fila de totales al final de la tabla."""
    n_rows      = len(df)
    n_cols      = len(df.columns)
    tot_row     = n_rows + 2
    col_names   = list(df.columns)
    importe_idx = col_names.index("importe")
    cliente_idx = col_names.index("cliente")

    totals_base = {
        "font_name": "Arial", "font_size": 10, "bold": True, "border": 1,
        "bg_color": TOTALS_BG, "font_color": TOTALS_FONT_COL, "valign": "vcenter"
    }
    totals_label_fmt = workbook.add_format({**totals_base, "align": "left"})
    totals_fmt       = workbook.add_format({**totals_base, "align": "center"})
    totals_num_fmt   = workbook.add_format({**totals_base, "align": "center",
                                            "num_format": "#,##0.00 €"})

    ws.write(tot_row, 0, "TOTALES", totals_label_fmt)
    ws.write(tot_row, cliente_idx, build_totals_summary(df), totals_fmt)
    ws.write(tot_row, importe_idx, df["importe"].sum(), totals_num_fmt)

    filled = {0, cliente_idx, importe_idx}
    for col_idx in range(n_cols):
        if col_idx not in filled:
            ws.write_blank(tot_row, col_idx, None, totals_fmt)

    ws.set_row(tot_row, 20)


# ------------------------------------------------------------------
# AUTOFIT COLUMNS
# ------------------------------------------------------------------

def autofit_columns(ws, df: pd.DataFrame) -> None:
    """Ajusta el ancho de columna al contenido máximo."""
    for col_idx, col_name in enumerate(df.columns):
        if col_name == "fecha_venta":
            width = 14
        else:
            try:
                max_len = max(
                    df[col_name].dropna().astype(str).map(len).max(),
                    len(col_name)
                )
                width = min(max_len + 2, 40)
            except (ValueError, TypeError):
                width = 12
        ws.set_column(col_idx, col_idx, width)


# ------------------------------------------------------------------
# WRITE REPORT SHEET
# ------------------------------------------------------------------

def write_report_sheet(workbook, df: pd.DataFrame) -> None:
    """
    Orquesta la escritura de la hoja Weekly_Report:
      1. Crear hoja y caché de formatos
      2. Título
      3. Cabecera
      4. Datos con formato condicional
      5. Totales
      6. Anchos de columna, freeze y filtro
    """
    ws        = workbook.add_worksheet("Weekly_Report")
    n_rows    = len(df)
    n_cols    = len(df.columns)
    fmt_cache = build_format_cache(workbook)

    _write_title(ws, workbook, n_cols)
    _write_header(ws, workbook, df)
    apply_column_formats(ws, df, fmt_cache)
    add_totals_row(ws, workbook, df)
    autofit_columns(ws, df)

    ws.freeze_panes(2, 0)
    ws.autofilter(1, 0, n_rows + 1, n_cols - 1)


# ------------------------------------------------------------------
# REPORT INFO SHEET
# ------------------------------------------------------------------

def write_info_sheet(workbook, df: pd.DataFrame, input_path: str) -> None:
    """Escribe la hoja Report_Info con metadatos del reporte."""
    ws = workbook.add_worksheet("Report_Info")

    label_fmt = workbook.add_format({
        "font_name": "Arial", "font_size": 10, "bold": True,
        "bg_color": "#D6E4F0", "border": 1, "align": "left"
    })
    value_fmt = workbook.add_format({
        "font_name": "Arial", "font_size": 10,
        "bg_color": "#EBF3FB", "border": 1, "align": "left"
    })

    cerradas = int((df["estado"] == "Cerrado").sum())
    cert     = int((df["certificacion"] != "Sin certificación").sum())

    meta = [
        ("Proyecto",            "excel-report-formatter"),
        ("Archivo fuente",      Path(input_path).name),
        ("Fecha generación",    datetime.now().strftime("%d/%m/%Y %H:%M")),
        ("Total filas",         str(len(df))),
        ("Ventas cerradas",     str(cerradas)),
        ("Ventas certificadas", str(cert)),
        ("Columnas",            str(len(df.columns))),
        ("Pipeline",            "excel-data-cleaner → excel-report-formatter"),
    ]

    ws.set_column(0, 0, 25)
    ws.set_column(1, 1, 45)

    for row_idx, (label, value) in enumerate(meta):
        ws.write(row_idx, 0, label, label_fmt)
        ws.write(row_idx, 1, value, value_fmt)
        ws.set_row(row_idx, 20)

    ws.hide_gridlines(2)


# ------------------------------------------------------------------
# MAIN ENTRY POINT
# ------------------------------------------------------------------

def run_formatter(input_path: str, output_path: str) -> None:
    """
    Punto de entrada principal.
    Carga el Excel limpio, genera el reporte semanal y lo guarda.

    Args:
        input_path  : ruta al Excel limpio (output del excel-data-cleaner)
        output_path : ruta donde se guarda el reporte semanal
    """
    df = load_clean_excel(input_path)
    validate_required_columns(df)

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    workbook = xlsxwriter.Workbook(output_path)

    write_report_sheet(workbook, df)
    write_info_sheet(workbook, df, input_path)

    workbook.close()
    print(f"  ✓ Reporte guardado: {output_path}")
