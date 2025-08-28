from __future__ import annotations
import argparse
from pathlib import Path
import pandas as pd
from decimal import Decimal, InvalidOperation
import math

# --- Args obligatorios (no hay defaults ni rutas fijas) ----------------------
def _parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(add_help=True, description="Stocktake simple (dos scanners + products).")
    p.add_argument("--scanner1", required=True, help="Ruta a scanner1 (xlsx/csv)")
    p.add_argument("--scanner2", required=True, help="Ruta a scanner2 (xlsx/csv)")
    p.add_argument("--products", required=True, help="Ruta a products (csv/xlsx)")
    p.add_argument("--outdir",   required=True, help="Directorio de salida")
    return p.parse_args()

# --- Heurísticas mínimas de columnas ----------------------------------------
BARCODE_CANDS = ["barcode", "bar code", "code", "ean", "upc", "codigo", "código"]
COUNT_CANDS   = ["count", "qty", "quantity", "cantidad", "scans"]

def _norm(s) -> str:
    if s is None:
        return ""
    return str(s).strip().lower().replace("\n", " ").replace("\r", " ")

def _find_col(df: pd.DataFrame, cands: list[str]) -> str | None:
    norm2real = {_norm(c): c for c in df.columns}
    # exactas
    for c in cands:
        if _norm(c) in norm2real:
            return norm2real[_norm(c)]
    # contains
    for real in df.columns:
        n = _norm(real)
        if any(_norm(c) in n for c in cands):
            return real
    return None

def _clean_barcode(x) -> str:
    # Normaliza a string de dígitos (soporta notación científica y sufijo '.0')
    if x is None:
        return ""
    if isinstance(x, float) and math.isnan(x):
        return ""
    s = str(x).strip()
    if s == "":
        return ""
    s = s.replace(",", "")
    try:
        if "e" in s.lower():
            d = Decimal(s)
            s = str(d.quantize(Decimal(1)))
    except InvalidOperation:
        pass
    if s.endswith(".0"):
        s = s[:-2]
    s = "".join(ch for ch in s if ch.isdigit())
    return s

def _load_scanner(path: Path) -> pd.DataFrame:
    # 1) Lectura inicial con header por defecto
    if path.suffix.lower() in (".xlsx", ".xls"):
        df = pd.read_excel(path)
    elif path.suffix.lower() == ".csv":
        df = pd.read_csv(path)
    else:
        raise ValueError(f"Formato no soportado para scanner: {path.suffix}")

    # Normalizar headers a str
    df.columns = [str(c).strip() if c is not None else "" for c in df.columns]

    # Si está vacío, devolvemos esquema
    if df.empty:
        return pd.DataFrame(columns=["barcode", "count"])

    # Intento normal (con headers)
    bcol = _find_col(df, BARCODE_CANDS)
    ccol = _find_col(df, COUNT_CANDS) if bcol else None

    # 2) Fallback: si NO encontramos la columna de barcode,
    # es probable que el archivo NO tenga encabezados y la primera fila haya quedado como header.
    if not bcol:
        # Releer sin encabezados
        if path.suffix.lower() in (".xlsx", ".xls"):
            df = pd.read_excel(path, header=None)
        else:  # csv
            df = pd.read_csv(path, header=None)

        # Asignar nombres mínimos según cantidad de columnas
        if df.shape[1] >= 2:
            df = df.iloc[:, :2].copy()
            df.columns = ["barcode", "count"]
        else:
            df.columns = ["barcode"]

        # Limpiar códigos
        df["barcode"] = df["barcode"].map(_clean_barcode)
        df = df[df["barcode"] != ""]

        # Conteo: si hay 'count' la normalizamos; si no, asumimos 1 por fila
        if "count" in df.columns:
            df["count"] = pd.to_numeric(df["count"], errors="coerce").fillna(0).astype(int)
            agg = df.groupby("barcode", as_index=False)["count"].sum()
        else:
            df["_u"] = 1
            agg = df.groupby("barcode", as_index=False)["_u"].sum().rename(columns={"_u": "count"})

        return agg

    # 3) Camino normal (sí encontramos barcode en la primera lectura)
    df[bcol] = df[bcol].map(_clean_barcode)
    df = df[df[bcol] != ""]
    if ccol:
        df[ccol] = pd.to_numeric(df[ccol], errors="coerce").fillna(0).astype(int)
        agg = df.groupby(bcol, as_index=False)[ccol].sum().rename(columns={bcol: "barcode", ccol: "count"})
    else:
        df["_u"] = 1
        agg = df.groupby(bcol, as_index=False)["_u"].sum().rename(columns={bcol: "barcode", "_u": "count"})

    return agg

def _load_products(path: Path) -> pd.DataFrame:
    # Leer como texto para preservar EAN largos y evitar notación científica
    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path, dtype=str)
    elif path.suffix.lower() in (".xlsx", ".xls"):
        df = pd.read_excel(path, dtype=str)
    else:
        raise ValueError(f"Formato no soportado para products: {path.suffix}")

    # Headers a str
    df.columns = [str(c).strip() if c is not None else "" for c in df.columns]
    if df.empty:
        raise ValueError("Products vacío.")

    # Detectar columnas clave
    bcol = _find_col(df, BARCODE_CANDS) or "Barcode"
    if bcol not in df.columns:
        raise ValueError("Products: falta columna de barcode (ej: 'Barcode').")

    name_cands = ["productname", "name", "product", "description", "descripcion"]
    id_cands   = ["productid", "id", "product id", "lightspeed id", "ls_id"]

    ncol = _find_col(df, name_cands)  # puede ser None
    icol = _find_col(df, id_cands)    # puede ser None

    # Construir un DF alineado fila a fila y recién después limpiar/deduplicar
    bar = df[bcol].map(_clean_barcode)
    pname = (df[ncol].astype(str).str.strip() if ncol else bar)
    pid = (df[icol].astype(str).str.strip() if icol else pd.Series(range(1, len(df) + 1), index=df.index))

    out = pd.DataFrame({
        "ProductID": pid,
        "ProductName": pname,
        "Barcode": bar,
    })

    # Quitar barcodes vacíos y deduplicar por Barcode manteniendo la primera ocurrencia
    out = out[out["Barcode"] != ""].drop_duplicates(subset=["Barcode"], keep="first").reset_index(drop=True)

    return out[["ProductID", "ProductName", "Barcode"]]


def _match(scans: pd.DataFrame, products: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    # Match exacto
    exact = scans.merge(products, left_on="barcode", right_on="Barcode", how="left")
    matched = exact.dropna(subset=["ProductID"])[["ProductID","ProductName","count"]]

    # Pendientes para sufijo
    pend = exact[exact["ProductID"].isna()][["barcode","count"]]
    if not pend.empty:
        pro = products.copy()
        rows=[]
        un=[]
        for _,r in pend.iterrows():
            b=str(r["barcode"]); c=int(r["count"])
            cand = pro[pro["Barcode"].astype(str).str.endswith(b)]
            if len(cand)==1:
                pr=cand.iloc[0]
                rows.append({"ProductID":pr["ProductID"],"ProductName":pr["ProductName"],"count":c})
            elif len(cand)>1:
                pr=cand.iloc[cand["Barcode"].str.len().argmax()]
                rows.append({"ProductID":pr["ProductID"],"ProductName":pr["ProductName"],"count":c})
            else:
                un.append({"scanned_barcode":b,"count":c})
        if rows:
            matched = pd.concat([matched, pd.DataFrame(rows)], ignore_index=True)
        unmatched = pd.DataFrame(un)
    else:
        unmatched = pd.DataFrame(columns=["scanned_barcode","count"])

    # Consolidar por producto
    if not matched.empty:
        matched = matched.groupby(["ProductID","ProductName"], as_index=False)["count"].sum()
        matched = matched.sort_values(["ProductName","ProductID"])
    return matched, unmatched

def main() -> int:
    args = _parse_args()
    sc1 = Path(args.scanner1); sc2 = Path(args.scanner2); pr = Path(args.products)
    outdir = Path(args.outdir); outdir.mkdir(parents=True, exist_ok=True)

    s1 = _load_scanner(sc1)
    s2 = _load_scanner(sc2)
    scans = pd.concat([s1,s2], ignore_index=True)
    scans = scans.groupby("barcode", as_index=False)["count"].sum()

    products = _load_products(pr)
    matched, unmatched = _match(scans, products)

    final_out = outdir / "final_count.csv"
    matched.to_csv(final_out, index=False)

    if not unmatched.empty:
        unmatched_out = outdir / "unmatched_barcodes.xlsx"
        # Asegurar que el barcode quede como TEXTO en Excel (evita 1.23E+13)
        unmatched = unmatched.copy()
        unmatched["scanned_barcode"] = unmatched["scanned_barcode"].astype(str)
        unmatched["count"] = pd.to_numeric(unmatched["count"], errors="coerce").fillna(0).astype(int)
        unmatched.to_excel(unmatched_out, index=False)  # requiere openpyxl instalado
        print(f"✅ final_count: {final_out}")
        print(f"📄 unmatched barcodes:   {unmatched_out}")
    else:
        print(f"✅ final_count: {final_out}")
        print("🎉 No unmatched barcodes.")

    return 0

if __name__ == "__main__":
    raise SystemExit(main())
