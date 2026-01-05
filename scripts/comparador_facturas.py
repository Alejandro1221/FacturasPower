from __future__ import annotations
from pathlib import Path
from decimal import Decimal
from typing import Optional, List, Dict, Any
import re
import pandas as pd


# Usamos el extractor que ya tienes
from extraer_TotalFactura import extraer_total, DEFAULT_DIR as DEFAULT_PDF_DIR

# ---------------- Utils ----------------
def _norm_amount_to_decimal(s: str | float | int | None) -> Optional[Decimal]:
    if s is None:
        return None
    s = str(s).strip()
    if not s or s.lower() in ("nan", "none"):
        return None
    s = s.replace(" ", "").replace(chr(160), "")
    # quita símbolos monetarios comunes
    s = re.sub(r"^[A-Za-z\$\s]*", "", s)
    # miles con punto, decimal con coma -> normalizamos
    s = s.replace(".", "").replace(",", ".")
    try:
        return Decimal(s)
    except Exception:
        return None

def _detect_col(df: pd.DataFrame, keywords: List[str]) -> Optional[str]:
    for c in df.columns:
        name = str(c).strip().lower()
        if any(k in name for k in keywords):
            return c
    return None

def _read_table_any(path: str) -> pd.DataFrame:
    low = path.lower()
    if low.endswith(".csv"):
        # Detectar separador por muestreo
        with open(path, "rb") as fh:
            head = fh.read(4096)
        sample = head.decode("latin-1", errors="ignore")
        sep_guess = ";" if sample.count(";") > sample.count(",") else ","

        # Probar encodings comunes
        for enc in ("utf-8-sig", "latin-1", "cp1252"):
            try:
                df = pd.read_csv(path, sep=sep_guess, dtype=str, keep_default_na=False, encoding=enc)
                break
            except UnicodeDecodeError:
                df = None
        if df is None:
            raise Exception("No se pudo leer el CSV (encodings: utf-8, latin-1, cp1252).")

    elif low.endswith((".xlsx", ".xlsm", ".xltx", ".xltm")):
        df = pd.read_excel(path, engine="openpyxl", dtype=str, keep_default_na=False)
    elif low.endswith(".xls"):
        df = pd.read_excel(path, engine="xlrd", dtype=str, keep_default_na=False)
    else:
        df = pd.read_excel(path, engine="openpyxl", dtype=str, keep_default_na=False)

    # Normalizar espacios en los nombres de columnas
    df.columns = [str(c).strip() for c in df.columns]
    return df

# ---------------- Core ----------------
def comparar_desde_excel(
    path_excel: str,
    carpeta_pdfs: str | Path | None = None,
    col_factura: str | None = None,
    col_total: str | None = None,
    limite: Optional[int] = None,
) -> List[Dict[str, Any]]:
    """
    Lee el Excel, localiza columnas de 'factura' y 'total',
    busca el PDF {factura}.pdf en carpeta_pdfs y compara totales.
    Retorna una lista de dicts con el resultado por factura.
    """
    df = _read_table_any(path_excel)

    # detectar columnas si no vienen forzadas
    if col_factura is None:
        col_factura = _detect_col(df, ["factura"])
    if col_total is None:
        col_total = _detect_col(df, ["total", "valor total", "total factura", "total a pagar"])

    if not col_factura or not col_total:
        raise ValueError(f"No se detectaron columnas. factura={col_factura!r}, total={col_total!r}")

    pdf_dir = Path(carpeta_pdfs or DEFAULT_PDF_DIR)
    pdf_dir.mkdir(parents=True, exist_ok=True)

    resultados: List[Dict[str, Any]] = []

    df[col_factura] = df[col_factura].astype(str).str.strip()
    df[col_total] = df[col_total].astype(str).str.strip()

    # Filtrar filas vacías o nulas
    df = df[df[col_factura].notna() & (df[col_factura] != "")]
    if df.empty:
        raise ValueError("El archivo no contiene facturas válidas (columna vacía).")

    rows = df[[col_factura, col_total]].values.tolist()
    if limite:
        rows = rows[:limite]

    for factura, total_excel_raw in rows:
        factura = str(factura).strip()
        total_excel = _norm_amount_to_decimal(total_excel_raw)

        if not factura:
            resultados.append({
                "factura": factura, "estado": "fila_sin_factura",
                "total_excel": total_excel, "total_pdf": None,
                "detalle": "Factura vacía en Excel"
            })
            continue

        pdf_path = pdf_dir / f"{factura}.pdf"
        if not pdf_path.exists():
            resultados.append({
                "factura": factura, "estado": "pdf_no_encontrado",
                "total_excel": total_excel, "total_pdf": None,
                "detalle": f"No existe {pdf_path.name}"
            })
            continue

        try:
            info = extraer_total(pdf_path)
            total_pdf = info.get("total")
            metodo = info.get("metodo", "?")
        except Exception as e:
            resultados.append({
                "factura": factura, "estado": "error_leyendo_pdf",
                "total_excel": total_excel, "total_pdf": None,
                "detalle": str(e)
            })
            continue

        if total_excel is None or total_pdf is None:
            resultados.append({
                "factura": factura, "estado": "dato_faltante",
                "total_excel": total_excel, "total_pdf": total_pdf,
                "detalle": "Falta total Excel o extracción PDF"
            })
            continue

        # comparación exacta en Decimal 
        if total_excel == total_pdf:
            estado = "OK"
        else:
            estado = "NO_COINCIDE"

        resultados.append({
            "factura": factura,
            "estado": estado,
            "total_excel": total_excel,
            "total_pdf": total_pdf,
            "detalle": f"metodo={metodo}"
        })

    return resultados


def imprimir_resumen(resultados: List[Dict[str, Any]]) -> None:
    ok = sum(1 for r in resultados if r["estado"] == "OK")
    fail = sum(1 for r in resultados if r["estado"] == "NO_COINCIDE")
    no_pdf = sum(1 for r in resultados if r["estado"] == "pdf_no_encontrado")
    faltante = sum(1 for r in resultados if r["estado"] in ("dato_faltante", "fila_sin_factura"))
    errores = sum(1 for r in resultados if r["estado"] == "error_leyendo_pdf")

    print("\n=== RESULTADO COMPARACIÓN ===")
    for r in resultados:
        print(f"[{r['estado']}] {r['factura']}: Excel={r['total_excel']} | PDF={r['total_pdf']} | {r['detalle']}")
    print("\n--- Resumen ---")
    print(f"  OK:            {ok}")
    print(f"  NO_COINCIDE:   {fail}")
    print(f"  pdf_no_encontrado: {no_pdf}")
    print(f"  dato/otros:    {faltante}")
    print(f"  errores_pdf:   {errores}")
    print(f"  TOTAL FILAS:   {len(resultados)}")

# Uso directo por consola 
if __name__ == "__main__":
    import argparse
    p = argparse.ArgumentParser(description="Comparar totales Excel vs PDF")
    p.add_argument("excel", help="Ruta del archivo Excel/CSV con columnas Factura y Total")
    p.add_argument("--pdfs", help="Carpeta de PDFs (por defecto: Facturas_descargadas)", default=str(DEFAULT_PDF_DIR))
    p.add_argument("--col-factura", default=None)
    p.add_argument("--col-total", default=None)
    p.add_argument("--limite", type=int, default=None)
    args = p.parse_args()

    res = comparar_desde_excel(args.excel, args.pdfs, args.col_factura, args.col_total, args.limite)
    imprimir_resumen(res)
