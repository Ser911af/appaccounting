# app_conciliacion.py
# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import unicodedata
from typing import Optional, List

# =========================
# Config & Título
# =========================
st.set_page_config(page_title="Conciliación de Cartera (Automática)", layout="wide")
st.title("Conciliación de Cartera: CierreCartera vs BALANCE 13452501 (Automática)")
st.caption("App fija para este archivo: detecta columnas y concilia por clave 'piso-num' (ej. 1-9803). Sin selectores.")

# =========================
# Utilidades
# =========================
def normalize_text(s: str) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    s = str(s)
    s = "".join(
        c for c in unicodedata.normalize("NFD", s)
        if unicodedata.category(c) != "Mn"
    )
    return s.lower().strip()

def make_unique(names: List[str]) -> List[str]:
    seen, out = {}, []
    for n in names:
        n = str(n)
        if n not in seen:
            seen[n] = 0
            out.append(n)
        else:
            seen[n] += 1
            out.append(f"{n}_{seen[n]}")
    return out

def build_table_from_row(df_raw: pd.DataFrame, header_row_idx: int) -> pd.DataFrame:
    headers = df_raw.iloc[header_row_idx].astype(str).tolist()
    headers = [
        h if normalize_text(h) not in ("", "unnamed: 0", "nan") else f"col_{i}"
        for i, h in enumerate(headers)
    ]
    headers = make_unique(headers)
    data = df_raw.iloc[header_row_idx+1:].copy()
    data.columns = headers
    data = data.dropna(how="all")
    for c in list(data.select_dtypes(include=["object"]).columns):
        data[c] = data[c].astype(str).str.strip()
    return data

def drop_all_empty_columns(df: pd.DataFrame) -> pd.DataFrame:
    tmp = df.copy()
    for c in tmp.columns:
        if tmp[c].dtype == object:
            tmp[c] = tmp[c].replace("", np.nan)
    mask_nonempty = tmp.notna().any(axis=0)
    return df.loc[:, mask_nonempty.values]

def to_amount(series: pd.Series) -> pd.Series:
    s = series.fillna("").astype(str)
    s = s.replace({r'[^0-9\-,\.]': ''}, regex=True)
    # ES: miles con punto, decimales con coma
    s = s.str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
    return pd.to_numeric(s, errors='coerce')

APT_SEP_REGEX = r"[-_/\.]"  # -, _, /, .
PI_SO_NUM_PATTERN = re.compile(rf"^\s*\d+\s*{APT_SEP_REGEX}\s*\d+\s*$")
PI_SO_NUM_SEARCH = re.compile(rf"(\d+)\s*{APT_SEP_REGEX}\s*(\d+)")

def normalize_apto_key(s: str) -> Optional[str]:
    if pd.isna(s):
        return None
    text = str(s).strip()
    m = PI_SO_NUM_SEARCH.search(text)
    if not m:
        return None
    piso = int(m.group(1))
    num  = int(m.group(2))
    return f"{piso}-{num}"

def find_col_fuzzy(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    cols = list(df.columns)
    ncols = [normalize_text(c) for c in cols]
    best, score_best = None, -1
    for i, nc in enumerate(ncols):
        score = sum(kw in nc for kw in candidates)
        if score > score_best:
            best, score_best = cols[i], score
    return best

def pick_apto_col_by_pattern(df: pd.DataFrame) -> Optional[str]:
    """
    Selecciona la mejor columna que contenga valores tipo 'piso-num' (1-9803),
    puntuando por cantidad de matches regex en las primeras filas no vacías.
    """
    best_col, best_hits = None, -1
    sample_n = min(200, len(df))
    for col in df.columns:
        s = df[col].dropna().astype(str).head(sample_n)
        hits = s.apply(lambda x: 1 if PI_SO_NUM_SEARCH.search(x) else 0).sum()
        if hits > best_hits:
            best_hits = hits
            best_col = col
    return best_col

def pick_amount_col(df: pd.DataFrame, prefer_keywords: List[str]) -> Optional[str]:
    """
    Elige columna de montos: primero por keywords, si empate/ausencia
    escoge la columna numérica con mayor suma absoluta.
    """
    kw_col = find_col_fuzzy(df, prefer_keywords)
    if kw_col is not None:
        return kw_col
    best_col, best_abs_sum = None, -1
    for col in df.columns:
        nums = to_amount(df[col])
        score = np.nansum(np.abs(nums.values))
        if np.isfinite(score) and score > best_abs_sum:
            best_abs_sum = score
            best_col = col
    return best_col

def synthesize_apto_from_block_and_code(
    df: pd.DataFrame,
    code_cols: List[str],
    block_cols: List[str],
    new_col_name: str = "_apto_sintetico"
) -> Optional[str]:
    """
    Crea una columna tipo 'piso-num' (1-9801) combinando:
    - código (ej: 'Inmueble Código' -> 9801)
    - bloque/piso (ej: 'Inmueble Bloque' -> 01 -> 1)
    """
    for code_col in code_cols:
        for block_col in block_cols:
            if code_col in df.columns and block_col in df.columns:
                bloque_raw = df[block_col].astype(str).str.extract(r"(\d+)")[0]
                codigo_raw = df[code_col].astype(str).str.extract(r"(\d{3,5})")[0]

                piso = bloque_raw.str.lstrip("0")
                piso = piso.replace("", np.nan)

                df[new_col_name] = np.where(
                    piso.notna() & codigo_raw.notna(),
                    piso + "-" + codigo_raw,
                    np.nan
                )

                hits = (
                    df[new_col_name]
                    .dropna()
                    .astype(str)
                    .apply(lambda x: 1 if PI_SO_NUM_SEARCH.search(x) else 0)
                    .sum()
                )
                if hits > 0:
                    return new_col_name

                df.drop(columns=[new_col_name], inplace=True, errors="ignore")

    return None

def col_has_piso_num_pattern(df: pd.DataFrame, col: Optional[str]) -> bool:
    if col is None or col not in df.columns:
        return False
    s = df[col].dropna().astype(str).head(300)
    hits = s.apply(lambda x: 1 if PI_SO_NUM_SEARCH.search(x) else 0).sum()
    return hits > 0

# =========================
# Parámetros fijos del caso
# =========================
SHEET_CIERRE = "Cierre"
SHEET_BALANCE = "Balance"
HDR_CIERRE_IDX = 7   # fila 8 de Excel
HDR_BALANCE_IDX = 2  # fila 3 de Excel
TOLERANCE_DEFAULT = 0.01

# =========================
# Carga de archivo
# =========================
uploaded = st.file_uploader(
    "Sube tu Excel (debe contener 'CierreCartera' y 'BALANCE 13452501')",
    type=["xlsx"]
)
if not uploaded:
    st.info("Sube el archivo para continuar. Este flujo está fijado a esas 2 hojas y filas de encabezado.")
    st.stop()

try:
    xls = pd.ExcelFile(uploaded)
    sheet_names = xls.sheet_names
    missing = [s for s in [SHEET_CIERRE, SHEET_BALANCE] if s not in sheet_names]
    if missing:
        st.error(
            f"No encuentro estas hojas requeridas: {', '.join(missing)}.\n"
            f"Hojas disponibles: {', '.join(sheet_names)}"
        )
        st.stop()
except Exception as e:
    st.error(f"No pude leer el Excel: {e}")
    st.stop()

# Lectura cruda y construcción con override de encabezados
raw1 = pd.read_excel(uploaded, sheet_name=SHEET_CIERRE, header=None, dtype=str)
raw2 = pd.read_excel(uploaded, sheet_name=SHEET_BALANCE, header=None, dtype=str)

df1 = build_table_from_row(raw1, HDR_CIERRE_IDX)
df2 = build_table_from_row(raw2, HDR_BALANCE_IDX)

df1 = drop_all_empty_columns(df1)
df2 = drop_all_empty_columns(df2)

st.success("Hojas cargadas con override fijo.")
with st.expander("Vista previa (Cierre)"):
    st.dataframe(df1.head(12), use_container_width=True)
with st.expander("Vista previa (Balance)"):
    st.dataframe(df2.head(12), use_container_width=True)

# =========================
# Detección automática de columnas
# =========================

# ---- Cierre: columna piso-num (apto) y valor cobro ----
apto_cierre_col = pick_apto_col_by_pattern(df1)
if apto_cierre_col is None:
    apto_cierre_col = find_col_fuzzy(
        df1,
        ["apto", "apart", "nro", "numero", "inmueble"]
    )

# Si no hay columna ya en formato 1-9803, sintetizar desde Inmueble Código + Inmueble Bloque
if not col_has_piso_num_pattern(df1, apto_cierre_col):
    synthesized = synthesize_apto_from_block_and_code(
        df1,
        code_cols=["Inmueble Código", "inmueble codigo", "codigo"],
        block_cols=["Inmueble Bloque", "bloque"],
        new_col_name="_apto_sintetico_cierre"
    )
    if synthesized is not None:
        apto_cierre_col = synthesized

valor_cobro_col = pick_amount_col(
    df1,
    ["valor cobro", "valor a cobrar", "valor cobrado", "cobro", "cuota", "facturado", "valor"]
)

# ---- Balance: apto (NIT), Nuevo Saldo y Cuenta (para 1345*) ----
apto_balance_col = pick_apto_col_by_pattern(df2)
if apto_balance_col is None:
    apto_balance_col = find_col_fuzzy(
        df2,
        ["nit", "nombre nit", "apto", "apart", "unidad", "inmueble"]
    )

# Monto: prioridad absoluta a "Nuevo Saldo"
nuevo_saldo_col = None
for col in df2.columns:
    if normalize_text(col) == "nuevo saldo":
        nuevo_saldo_col = col
        break

if nuevo_saldo_col is None:
    nuevo_saldo_col = pick_amount_col(
        df2,
        ["nuevo saldo", "saldo nuevo", "saldo final", "saldo", "balance", "cartera", "deuda"]
    )

# Columna de número de cuenta (para identificar las 1345***)
cuenta_balance_col = find_col_fuzzy(
    df2,
    ["cuenta", "codigo", "código", "cod"]
)

chosen_core = {
    "apto_cierre_col": apto_cierre_col,
    "valor_cobro_col": valor_cobro_col,
    "apto_balance_col": apto_balance_col,
    "nuevo_saldo_col": nuevo_saldo_col
}
chosen_extended = {
    **chosen_core,
    "cuenta_balance_col": cuenta_balance_col
}

if any(chosen_core[k] is None for k in chosen_core):
    st.error(f"No pude detectar todas las columnas necesarias automáticamente.\nDetecciones: {chosen_extended}")
    st.stop()

st.info(
    f"Columnas detectadas automáticamente:\n"
    f"- Cierre (apto): **{apto_cierre_col}**\n"
    f"- Cierre (valor cobro): **{valor_cobro_col}**\n"
    f"- Balance (apto): **{apto_balance_col}**\n"
    f"- Balance (Nuevo Saldo): **{nuevo_saldo_col}**\n"
    f"- Balance (Cuenta para 1345*): **{cuenta_balance_col}**"
)

# =========================
# Conciliación automática
# =========================
tolerance = TOLERANCE_DEFAULT

# Montos numéricos
df1["_valor_cobro_num"] = to_amount(df1[valor_cobro_col])
df2["_nuevo_saldo_num"] = to_amount(df2[nuevo_saldo_col])

# Flag de cuenta 1345* en Balance
if cuenta_balance_col is not None and cuenta_balance_col in df2.columns:
    df2["_is_1345"] = (
        df2[cuenta_balance_col]
        .astype(str)
        .str.strip()
        .str.startswith("1345")
    )
else:
    df2["_is_1345"] = False

# Monto solo para cuentas 1345*
df2["_nuevo_saldo_1345_num"] = np.where(
    df2["_is_1345"],
    df2["_nuevo_saldo_num"],
    0.0
)

# =========================
# Clave piso-num con filtrado de valores no válidos
# =========================
def is_valid_apto_format(s: str) -> bool:
    """Acepta solo formatos tipo 1-9803, 2-9901, etc. (no P205, no vacíos, no 999.999-1)"""
    if pd.isna(s):
        return False
    text = str(s).strip()
    return bool(re.match(r"^\d{1,2}[-_/\.]\d{3,5}$", text))

# Normalización y filtros
df1["_apto_key"] = df1[apto_cierre_col].apply(normalize_apto_key)
df2["_apto_key"] = df2[apto_balance_col].apply(normalize_apto_key)

df1 = df1[df1[apto_cierre_col].apply(is_valid_apto_format)]
df2 = df2[df2[apto_balance_col].apply(is_valid_apto_format)]

# =========================
# Agregaciones
# =========================
g1 = (
    df1.dropna(subset=["_apto_key"])
       .groupby("_apto_key", as_index=False)
       .agg(
           valor_cobro_sum=("_valor_cobro_num", "sum"),
           conteo_registros=(apto_cierre_col, "count")
       )
)

# Balance: total y 1345* por apartamento
g2 = (
    df2.dropna(subset=["_apto_key"])
       .groupby("_apto_key", as_index=False)
       .agg(
           nuevo_saldo_sum=("_nuevo_saldo_num", "sum"),
           nuevo_saldo_1345_sum=("_nuevo_saldo_1345_num", "sum")
       )
)

# Join y diferencia (contra el total de Nuevo Saldo)
res = pd.merge(g1, g2, on="_apto_key", how="outer")
res["valor_cobro_sum"] = res["valor_cobro_sum"].fillna(0.0)
res["nuevo_saldo_sum"] = res["nuevo_saldo_sum"].fillna(0.0)
res["nuevo_saldo_1345_sum"] = res["nuevo_saldo_1345_sum"].fillna(0.0)

res["diferencia"] = res["valor_cobro_sum"] - res["nuevo_saldo_sum"]

# Filtrar diferencias ≠ 0 (tolerancia)
conciliacion = res[res["diferencia"].abs() > tolerance].sort_values("_apto_key")

# =========================
# Resultados y descarga
# =========================
st.markdown("### Resultados")
m1, m2, m3, m4 = st.columns(4)
with m1:
    st.metric("Aptos en Cierre", int(g1.shape[0]))
with m2:
    st.metric("Aptos en Balance", int(g2.shape[0]))
with m3:
    st.metric("Coincidencias (outer join)", int(res.shape[0]))
with m4:
    st.metric("Diferencias ≠ 0", int(conciliacion.shape[0]))

tabs = st.tabs(["Conciliación", "Match Total", "Agregado Cierre", "Agregado Balance"])
with tabs[0]:
    st.dataframe(conciliacion.reset_index(drop=True), use_container_width=True)
with tabs[1]:
    st.dataframe(res.sort_values("_apto_key").reset_index(drop=True), use_container_width=True)
with tabs[2]:
    st.dataframe(g1.sort_values("_apto_key").reset_index(drop=True), use_container_width=True)
with tabs[3]:
    st.dataframe(g2.sort_values("_apto_key").reset_index(drop=True), use_container_width=True)

def build_output_excel() -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as w:
        g1.rename(columns={"_apto_key": "apto_key"}).to_excel(w, "agregado_cierre", index=False)
        g2.rename(columns={"_apto_key": "apto_key"}).to_excel(w, "agregado_balance", index=False)
        res.rename(columns={"_apto_key": "apto_key"}).to_excel(w, "match_total", index=False)
        conciliacion.rename(columns={"_apto_key": "apto_key"}).to_excel(w, "conciliacion", index=False)
    return output.getvalue()

st.markdown("### Descargar resultados")
st.download_button(
    "Descargar Excel (agregados, match y conciliación)",
    data=build_output_excel(),
    file_name="conciliacion_cartera.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

with st.expander("Diagnóstico (columnas detectadas)"):
    st.json(chosen_extended)

st.caption("Automático, con cuentas 1345 sumadas por apto. Si algo no cuadra, revisamos al 803… o a contabilidad 😉.")
