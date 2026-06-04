"""
Normalizador unificado de extractos bancarios
Seleccionás los archivos de cada banco y genera un Excel con:
  - TABLA GENERAL - Egresos
  - TABLA GENERAL - Ingresos
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import unicodedata
import re
from io import StringIO


# ─── helpers compartidos ─────────────────────────────────────────────────────

def _strip_accents(text: str) -> str:
    nfkd = unicodedata.normalize("NFD", text)
    return "".join(c for c in nfkd if unicodedata.category(c) != "Mn")


def _norm_colname(name: str) -> str:
    s = _strip_accents(str(name).strip()).lower()
    s = re.sub(r"[.\s]+", "_", s)
    s = re.sub(r"[^a-z0-9_]", "", s)
    return s.strip("_")


def _parse_arg_number(val) -> float | None:
    """Convierte formato argentino: (1.234,56) → -1234.56  /  1.234,56 → 1234.56"""
    if pd.isna(val):
        return None
    s = str(val).strip()
    if not s:
        return None
    negative = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    s = s.replace(".", "").replace(",", ".")
    try:
        num = float(s)
        return -num if negative else num
    except ValueError:
        return None


def _to_float(val) -> float | None:
    if pd.isna(val):
        return None
    try:
        return float(str(val).strip().replace(",", "."))
    except ValueError:
        return None


def _format_fecha(val) -> str:
    if pd.isna(val):
        return ""
    try:
        return pd.to_datetime(val).strftime("%d/%m/%Y")
    except Exception:
        return str(val)


def _find_col(df: pd.DataFrame, *keywords) -> str | None:
    for kw in keywords:
        for c in df.columns:
            if kw in c:
                return c
    return None


# ─── SANTANDER ───────────────────────────────────────────────────────────────

def _load_santander(filepath: str) -> pd.DataFrame:
    """
    El extracto Santander es un TSV disfrazado de .xls (o .csv/.txt).
    Busca la fila de encabezados por columnas 'fecha' + 'importe'.
    """
    ext = filepath.rsplit(".", 1)[-1].lower()

    if ext == "xlsx":
        df_raw = pd.read_excel(filepath, header=None, dtype=str)
        for i, row in df_raw.iterrows():
            cols = [_norm_colname(str(v)) for v in row.values]
            if "fecha" in cols and any("importe" in c or "monto" in c for c in cols):
                df_raw.columns = [_norm_colname(str(v)) for v in df_raw.iloc[i].values]
                return df_raw.iloc[i + 1:].copy().reset_index(drop=True)
        raise ValueError("No se encontró encabezado válido en el Excel de Santander.")

    # archivo de texto (xls falso, csv, txt, tsv)
    for enc in ("latin-1", "utf-8", "cp1252"):
        try:
            with open(filepath, "r", encoding=enc) as f:
                lines = f.readlines()
            break
        except (UnicodeDecodeError, LookupError):
            continue
    else:
        raise ValueError("No se pudo leer el archivo Santander.")

    header_idx = None
    for i, line in enumerate(lines):
        cols = [_norm_colname(c) for c in re.split(r"\t|;|,", line)]
        if "fecha" in cols and any("importe" in c or "monto" in c for c in cols):
            header_idx = i
            break
    if header_idx is None:
        raise ValueError("No se encontró fila de encabezados en el archivo Santander.")

    sample = lines[header_idx]
    sep = "\t" if "\t" in sample else (";" if ";" in sample else ",")
    df = pd.read_csv(StringIO("".join(lines[header_idx:])), sep=sep, dtype=str)
    df.columns = [_norm_colname(c) for c in df.columns]
    df = df.dropna(how="all")
    df = df[df.apply(lambda r: r.astype(str).str.strip().ne("").any(), axis=1)]
    return df


def _normalize_santander(df: pd.DataFrame):
    col_fecha    = _find_col(df, "fecha")
    col_importe  = _find_col(df, "importe", "monto")
    col_concepto = _find_col(df, "concepto", "detalle", "descripcion")

    if not col_fecha:
        raise ValueError("Santander: columna 'fecha' no encontrada.")
    if not col_importe:
        raise ValueError("Santander: columna 'importe'/'monto' no encontrada.")

    df = df.copy()
    df["_importe"] = df[col_importe].apply(_parse_arg_number)
    df = df.dropna(subset=["_importe"]).reset_index(drop=True)

    mask_eg = df["_importe"] < 0
    mask_in = df["_importe"] > 0

    eg = df[mask_eg].reset_index(drop=True)
    df_eg = pd.DataFrame({
        "BANCO":    "SANTANDER",
        "FECHA":    eg[col_fecha],
        "DATA 1":   eg[col_concepto] if col_concepto else "",
        "DATA 2":   "",
        "TOTAL":    eg["_importe"].abs(),
        "EGRESOS":  "",
        "RUBRO":    "",
        "SUBRUBRO": "",
        "PAGO":     "",
        "OBSERVAC": "",
    }).reset_index(drop=True)

    ing = df[mask_in].reset_index(drop=True)
    df_in = pd.DataFrame({
        "BANCO":    "SANTANDER",
        "FECHA":    ing[col_fecha],
        "DATA 1":   ing[col_concepto] if col_concepto else "",
        "DATA 2":   "",
        "TOTAL":    ing["_importe"],
        "INGRESOS": "",
        "RUBRO":    "",
        "SUBRUBRO": "",
        "ORIGEN":   "",
    }).reset_index(drop=True)

    return df_eg, df_in


# ─── GALICIA ─────────────────────────────────────────────────────────────────

def _load_galicia(filepath: str) -> pd.DataFrame:
    ext = filepath.rsplit(".", 1)[-1].lower()
    if ext in ("xlsx", "xls"):
        df = pd.read_excel(filepath, dtype=str)
    elif ext == "csv":
        for sep in (",", ";", "\t"):
            try:
                df = pd.read_csv(filepath, sep=sep, dtype=str, encoding="latin-1")
                if len(df.columns) > 2:
                    break
            except Exception:
                continue
    else:
        raise ValueError(f"Galicia: formato no soportado .{ext}")

    df.columns = [_norm_colname(c) for c in df.columns]

    missing = {"fecha", "debitos", "creditos"} - set(df.columns)
    if missing:
        raise ValueError(f"Galicia: columnas faltantes {missing}. Encontradas: {list(df.columns)}")

    return df


def _normalize_galicia(df: pd.DataFrame, banco: str = "GALICIA"):
    col_descripcion = _find_col(df, "descripcion", "descripci")
    col_grupo       = _find_col(df, "grupo")

    df = df.copy()
    df["_debitos"]  = df["debitos"].apply(_to_float).fillna(0)
    df["_creditos"] = df["creditos"].apply(_to_float).fillna(0)
    df["_fecha"]    = df["fecha"].apply(_format_fecha)

    mask_eg = df["_debitos"] > 0
    mask_in = df["_creditos"] > 0

    eg = df[mask_eg].reset_index(drop=True)
    df_eg = pd.DataFrame({
        "BANCO":    banco,
        "FECHA":    eg["_fecha"],
        "DATA 1":   eg[col_descripcion] if col_descripcion else "",
        "DATA 2":   eg[col_grupo]       if col_grupo       else "",
        "TOTAL":    eg["_debitos"],
        "EGRESOS":  "",
        "RUBRO":    "",
        "SUBRUBRO": "",
        "PAGO":     "",
        "OBSERVAC": "",
    }).reset_index(drop=True)

    ing = df[mask_in].reset_index(drop=True)
    df_in = pd.DataFrame({
        "BANCO":    banco,
        "FECHA":    ing["_fecha"],
        "DATA 1":   ing[col_descripcion] if col_descripcion else "",
        "DATA 2":   ing[col_grupo]       if col_grupo       else "",
        "TOTAL":    ing["_creditos"],
        "INGRESOS": "",
        "RUBRO":    "",
        "SUBRUBRO": "",
        "ORIGEN":   "",
    }).reset_index(drop=True)

    return df_eg, df_in


# ─── BANCO SAN JUAN ──────────────────────────────────────────────────────────

def _load_san_juan(filepath: str) -> pd.DataFrame:
    """
    .xls real (Excel 97). Tiene filas de metadata antes del header.
    Busca la fila con 'fecha' + 'monto' para usarla como encabezado.
    """
    ext = filepath.rsplit(".", 1)[-1].lower()
    engine = "xlrd" if ext == "xls" else None

    df_raw = pd.read_excel(filepath, header=None, dtype=str, engine=engine)

    header_idx = None
    for i, row in df_raw.iterrows():
        cols = [_norm_colname(str(v)) for v in row.values]
        if "fecha" in cols and any("monto" in c or "importe" in c for c in cols):
            header_idx = i
            break
    if header_idx is None:
        raise ValueError("San Juan: no se encontró fila de encabezados con 'Fecha' y 'Monto'.")

    df_raw.columns = [_norm_colname(str(v)) for v in df_raw.iloc[header_idx].values]
    df = df_raw.iloc[header_idx + 1:].copy().reset_index(drop=True)
    df = df.dropna(how="all")
    return df


def _normalize_san_juan(df: pd.DataFrame):
    col_fecha    = _find_col(df, "fecha")
    col_monto    = _find_col(df, "monto", "importe")
    col_concepto = _find_col(df, "concepto", "descripcion", "detalle")

    if not col_fecha:
        raise ValueError("San Juan: columna 'fecha' no encontrada.")
    if not col_monto:
        raise ValueError("San Juan: columna 'monto'/'importe' no encontrada.")

    df = df.copy()
    df["_monto"] = df[col_monto].apply(_to_float)
    df = df.dropna(subset=["_monto"]).reset_index(drop=True)

    mask_eg = df["_monto"] < 0
    mask_in = df["_monto"] > 0

    eg = df[mask_eg].reset_index(drop=True)
    df_eg = pd.DataFrame({
        "BANCO":    "SAN JUAN",
        "FECHA":    eg[col_fecha],
        "DATA 1":   eg[col_concepto] if col_concepto else "",
        "DATA 2":   "",
        "TOTAL":    eg["_monto"].abs(),
        "EGRESOS":  "",
        "RUBRO":    "",
        "SUBRUBRO": "",
        "PAGO":     "",
        "OBSERVAC": "",
    }).reset_index(drop=True)

    ing = df[mask_in].reset_index(drop=True)
    df_in = pd.DataFrame({
        "BANCO":    "SAN JUAN",
        "FECHA":    ing[col_fecha],
        "DATA 1":   ing[col_concepto] if col_concepto else "",
        "DATA 2":   "",
        "TOTAL":    ing["_monto"],
        "INGRESOS": "",
        "RUBRO":    "",
        "SUBRUBRO": "",
        "ORIGEN":   "",
    }).reset_index(drop=True)

    return df_eg, df_in


# ─── ICBC ────────────────────────────────────────────────────────────────────

def _load_icbc(filepath: str) -> pd.DataFrame:
    """
    CSV con separador ';', primera fila es metadata (nombre de cuenta).
    Header real en la segunda fila.
    Números con coma decimal: -319,92
    """
    ext = filepath.rsplit(".", 1)[-1].lower()

    if ext in ("xlsx", "xls"):
        engine = "xlrd" if ext == "xls" else None
        df_raw = pd.read_excel(filepath, header=None, dtype=str, engine=engine)
        for i, row in df_raw.iterrows():
            cols = [_norm_colname(str(v)) for v in row.values]
            if "fecha_contable" in cols or ("fecha" in cols and any("debito" in c for c in cols)):
                df_raw.columns = [_norm_colname(str(v)) for v in df_raw.iloc[i].values]
                df = df_raw.iloc[i + 1:].copy().reset_index(drop=True)
                return df.dropna(how="all")
        raise ValueError("ICBC: no se encontró fila de encabezados.")

    # CSV / TXT
    for enc in ("latin-1", "utf-8", "cp1252"):
        try:
            with open(filepath, "r", encoding=enc) as f:
                lines = f.readlines()
            break
        except (UnicodeDecodeError, LookupError):
            continue
    else:
        raise ValueError("ICBC: no se pudo leer el archivo.")

    # buscar fila de encabezados
    header_idx = None
    for i, line in enumerate(lines):
        cols = [_norm_colname(c) for c in re.split(r";|,|\t", line)]
        if "fecha_contable" in cols or ("fecha" in cols and any("debito" in c for c in cols)):
            header_idx = i
            break
    if header_idx is None:
        raise ValueError("ICBC: no se encontró fila de encabezados con 'Fecha contable' y 'Debito'.")

    df = pd.read_csv(
        StringIO("".join(lines[header_idx:])),
        sep=";", dtype=str, encoding="utf-8"
    )
    df.columns = [_norm_colname(c) for c in df.columns]
    return df.dropna(how="all")


def _normalize_icbc(df: pd.DataFrame):
    col_fecha    = _find_col(df, "fecha_contable", "fecha")
    col_debito   = _find_col(df, "debito")
    col_credito  = _find_col(df, "credito")
    col_concepto = _find_col(df, "concepto", "descripcion", "detalle")

    if not col_fecha:
        raise ValueError("ICBC: columna 'Fecha contable' no encontrada.")
    if not col_debito:
        raise ValueError("ICBC: columna 'Debito en $' no encontrada.")

    def _parse(val):
        if pd.isna(val) or str(val).strip() == "":
            return None
        s = str(val).strip().replace(".", "").replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return None

    df = df.copy()
    df["_debito"]  = df[col_debito].apply(_parse)
    df["_credito"] = df[col_credito].apply(_parse) if col_credito else 0

    # débitos: filas con valor en columna debito (abs para TOTAL)
    mask_eg = df["_debito"].notna() & (df["_debito"] != 0)
    # créditos: filas con valor en columna credito
    mask_in = df["_credito"].notna() & (df["_credito"] != 0)

    eg = df[mask_eg].reset_index(drop=True)
    df_eg = pd.DataFrame({
        "BANCO":    "ICBC",
        "FECHA":    eg[col_fecha],
        "DATA 1":   eg[col_concepto] if col_concepto else "",
        "DATA 2":   "",
        "TOTAL":    eg["_debito"].abs(),
        "EGRESOS":  "",
        "RUBRO":    "",
        "SUBRUBRO": "",
        "PAGO":     "",
        "OBSERVAC": "",
    }).reset_index(drop=True)

    ing = df[mask_in].reset_index(drop=True)
    df_in = pd.DataFrame({
        "BANCO":    "ICBC",
        "FECHA":    ing[col_fecha],
        "DATA 1":   ing[col_concepto] if col_concepto else "",
        "DATA 2":   "",
        "TOTAL":    ing["_credito"].abs(),
        "INGRESOS": "",
        "RUBRO":    "",
        "SUBRUBRO": "",
        "ORIGEN":   "",
    }).reset_index(drop=True)

    return df_eg, df_in


# ─── registro de bancos ──────────────────────────────────────────────────────

BANCOS = [
    {
        "nombre":    "Santander",
        "formatos":  "*.xls *.xlsx *.csv *.txt *.tsv",
        "load":      _load_santander,
        "normalize": _normalize_santander,
    },
    {
        "nombre":    "Galicia",
        "formatos":  "*.xls *.xlsx *.csv",
        "load":      _load_galicia,
        "normalize": _normalize_galicia,
    },
    {
        "nombre":    "Galicia Mas",
        "formatos":  "*.xls *.xlsx *.csv",
        "load":      _load_galicia,
        "normalize": lambda df: _normalize_galicia(df, banco="GALICIA MAS"),
    },
    {
        "nombre":    "San Juan",
        "formatos":  "*.xls *.xlsx *.csv",
        "load":      _load_san_juan,
        "normalize": _normalize_san_juan,
    },
    {
        "nombre":    "ICBC",
        "formatos":  "*.csv *.xls *.xlsx *.txt",
        "load":      _load_icbc,
        "normalize": _normalize_icbc,
    },
]


# ─── main ────────────────────────────────────────────────────────────────────

def main():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    todos_egresos  = []
    todos_ingresos = []
    resumen        = []

    for banco in BANCOS:
        nombre   = banco["nombre"]
        formatos = banco["formatos"]

        respuesta = messagebox.askyesno(
            f"Banco {nombre}",
            f"¿Querés cargar el extracto del Banco {nombre}?",
        )
        if not respuesta:
            continue

        filepath = filedialog.askopenfilename(
            title=f"Seleccionar extracto Banco {nombre}",
            filetypes=[
                (f"Formatos {nombre}", formatos),
                ("Todos los archivos", "*.*"),
            ],
        )
        if not filepath:
            messagebox.showwarning(f"Banco {nombre}", "No se seleccionó archivo. Se omite este banco.")
            continue

        try:
            df_raw = banco["load"](filepath)
            df_eg, df_in = banco["normalize"](df_raw)
            todos_egresos.append(df_eg)
            todos_ingresos.append(df_in)
            resumen.append(f"  {nombre}: {len(df_eg)} egresos, {len(df_in)} ingresos")
        except Exception as exc:
            messagebox.showerror(f"Error – Banco {nombre}", str(exc))
            continue

    if not todos_egresos and not todos_ingresos:
        messagebox.showwarning("Sin datos", "No se procesó ningún banco.")
        return

    df_egresos  = pd.concat(todos_egresos,  ignore_index=True) if todos_egresos  else pd.DataFrame()
    df_ingresos = pd.concat(todos_ingresos, ignore_index=True) if todos_ingresos else pd.DataFrame()

    out_path = filedialog.asksaveasfilename(
        title="Guardar reporte normalizado",
        defaultextension=".xlsx",
        filetypes=[("Excel", "*.xlsx")],
        initialfile="reporte banco normalizado.xlsx",
    )
    if not out_path:
        messagebox.showwarning("Cancelado", "No se guardó el archivo.")
        return

    try:
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            if not df_egresos.empty:
                df_egresos.to_excel(writer, sheet_name="TABLA GENERAL - Egresos",  index=False)
            if not df_ingresos.empty:
                df_ingresos.to_excel(writer, sheet_name="TABLA GENERAL - Ingresos", index=False)
    except Exception as exc:
        messagebox.showerror("Error al guardar", str(exc))
        return

    messagebox.showinfo(
        "Listo",
        f"Archivo generado:\n{out_path}\n\n"
        + "\n".join(resumen)
        + f"\n\n  Total egresos:  {len(df_egresos)} filas"
        + f"\n  Total ingresos: {len(df_ingresos)} filas",
    )


if __name__ == "__main__":
    main()
