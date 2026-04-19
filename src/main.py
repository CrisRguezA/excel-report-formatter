"""
ES:
Punto de entrada del pipeline excel-report-formatter.
Orquesta la carga, validación y generación del reporte semanal.

Uso:
    python src/main.py
    python src/main.py --input data/raw/mi_archivo.xlsx
    python src/main.py --input data/raw/mi_archivo.xlsx --output outputs/mi_reporte.xlsx

EN:
Entry point for the excel-report-formatter pipeline.
Orchestrates loading, validation, and weekly report generation.

Usage:
    python src/main.py
    python src/main.py --input data/raw/my_file.xlsx
    python src/main.py --input data/raw/my_file.xlsx --output outputs/my_report.xlsx
"""

# ------------------------------------------------------------------
# IMPORTS
# ------------------------------------------------------------------

import argparse
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))
from formatter import load_clean_excel, validate_required_columns, run_formatter


# ------------------------------------------------------------------
# DEFAULT PATHS
# ------------------------------------------------------------------

ROOT = Path(__file__).resolve().parent.parent
INPUT_FILE = ROOT / "data" / "raw" / "cleaned_sales_report.xlsx"
OUTPUT_FILE = ROOT / "outputs" / "weekly_sales_report.xlsx"


# ------------------------------------------------------------------
# SUMMARY HELPERS
# ------------------------------------------------------------------

def print_summary(df: pd.DataFrame, output_path: Path) -> None:
    """Imprime un resumen final del reporte generado."""
    total = len(df)
    cerradas = int((df["estado"] == "Cerrado").sum())
    certificadas = int((df["certificacion"] != "Sin certificación").sum())

    pct_cerradas = (cerradas / total * 100) if total > 0 else 0
    pct_certificadas = (certificadas / total * 100) if total > 0 else 0

    print("\nResumen:")
    print(f"  Filas procesadas    : {total}")
    print(f"  Ventas cerradas     : {cerradas} ({pct_cerradas:.0f}%)")
    print(f"  Ventas certificadas : {certificadas} ({pct_certificadas:.0f}%)")
    print(f"  Output              : {output_path}")


# ------------------------------------------------------------------
# PIPELINE
# ------------------------------------------------------------------

def run_pipeline(input_path: Path, output_path: Path) -> None:
    """Orquesta la ejecución del formatter con salida estructurada por consola."""
    print("\n" + "=" * 42)
    print("  Excel Report Formatter")
    print("=" * 42)

    print("\n[1] Cargando Excel limpio...")
    df = load_clean_excel(str(input_path))

    print("\n[2] Validando columnas...")
    validate_required_columns(df)
    print("    ✓ Esquema correcto")

    print("\n[3] Generando reporte...")
    run_formatter(str(input_path), str(output_path))
    print(f"    ✓ Reporte guardado: {output_path}")

    print_summary(df, output_path)
    print()


# ------------------------------------------------------------------
# CLI
# ------------------------------------------------------------------

def parse_args() -> argparse.Namespace:
    """Parsea argumentos de línea de comandos."""
    parser = argparse.ArgumentParser(
        description="Genera un reporte semanal Excel a partir de un dataset limpio."
    )
    parser.add_argument(
        "--input",
        type=Path,
        default=INPUT_FILE,
        help="Ruta del Excel limpio (output del excel-data-cleaner)",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=OUTPUT_FILE,
        help="Ruta del reporte Excel de salida",
    )
    return parser.parse_args()


# ------------------------------------------------------------------
# ENTRY POINT
# ------------------------------------------------------------------

if __name__ == "__main__":
    args = parse_args()
    try:
        run_pipeline(args.input, args.output)
    except Exception as e:
        print(f"\n✗ Error: {e}")
        sys.exit(1)