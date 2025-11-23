# app_conciliacion_cierre_balance.py
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
st.set_page_config(page_title="Conciliación de Cartera (Cierre vs Balance)", layout="wide")
st.title("Conciliación de Cartera: Cierre vs Balance (2 archivos)")
st.caption("Sube el archivo de Cierre y el de Balance. Se concilia por clave de apartamento y cuentas 1345*.")

# =========================
# Utilidades
# =========================
def normalize_text(s: str) -> str:
    """Normaliza texto: minúsculas, sin tildes, sin espacios extremos."""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    s = str(s)
    s = "".join(
        c for c in unicodedata.normalize("NFD", s)
        if unicodedata.category(c) != "Mn"
    )
    return s.lower().strip()

def drop_all_empty_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Elimina columnas totalmente vacías ('' o NaN)."""
    tmp = df.copy()
    for c in tmp.columns:
        if tmp[c].dtype == object:
            tmp[c] = tmp[c].replace("", np.nan)
    mask_nonempty = tmp.notna().any(axis=0)
    return df.loc[:, mask_nonempty.values]

def to_amount(series: pd.Series) -> pd.Series:
    """
    Convierte texto numérico en número.
    Soporta formato tipo 1.234.567,89 (punto miles, coma decimales).
    """
    s = series.fillna("").astype(str)
    s = s.replace({r'[^0-9\-,\.]': ''}, regex=True)
    s = s.str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
    return pd.to_numeric(s, errors='coerce')

def build_table_from_row(df_raw: pd.DataFrame, header_row_idx: int) -> pd.DataFrame:
    """
    Usa la fila `header_row_idx` como encabezados y devuelve el resto como datos.
    """
    headers = df_raw.iloc[header_row_idx].astype(str).tolist()
    headers_norm = []
    for i, h in enumerate(headers):
        h_norm = normalize_text(h)
        if h_norm in ("", "unnamed: 0", "nan"):
            headers_norm.append(f"col_{i}")
        else:
            headers_norm.append(h)

    # Asegurar unicidad
    seen, unique_headers = {}, []
    for h in headers_norm:
        if h not in seen:
            seen[h] = 0
            unique_headers.append(h)
        else:
            seen[h] += 1
            unique_headers.append(f"{h}_{seen[h]}")

    data = df_raw.iloc[header_row_idx+1:].copy()
    data.columns = unique_headers
    data = data.dropna(how="all")
    for c in list(data.select_dtypes(include=["object"]).columns):
        data[c] = data[c].astype(str).str.strip()
    return data

APT_SEP_REGEX = r"[-_/\.]"
PI_SO_NUM_SEARCH = re.compile(rf"(\d+)\s*{APT_SEP_REGEX}\s*(\d+)")

def normalize_apto_key(s: str) -> Optional[str]:
    """Extrae y normaliza clave tipo piso-num (1-9801) de un string."""
    if pd.isna(s):
        return None
    text = str(s).strip()
    m = PI_SO_NUM_SEARCH.search(text)
    if not m:
        return None
    piso = int(m.group(1))
    num  = int(m.group(2))
    return f"{piso}-{num}"

def is_valid_apto_format(s: str) -> bool:
    """Acepta solo formatos tipo 1-9803, 2-9901, etc."""
    if pd.isna(s):
        return False
    text = str(s).strip()
    return bool(re.match(r"^\d{1,2}[-_/\.]\d{3,5}$", text))

def find_col_fuzzy(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    """Encuentra la columna cuyo nombre se parece más a la lista de keywords."""
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

# =========================
# 1. Carga de archivos
# =========================
st.subheader("1. Carga de archivos")

cierre_file = st.file_uploader("📥 Archivo de Cierre", type=["xlsx"], key="cierre")
balance_file = st.file_uploader("📥 Archivo de Balance", type=["xlsx"], key="balance")

if not cierre_file or not balance_file:
    st.info("Sube ambos archivos (Cierre y Balance) para continuar.")
    st.stop()

# ---- Cierre: leer sin encabezado y seleccionar fila de header ----
try:
    raw_cierre = pd.read_excel(cierre_file, header=None, dtype=str)
except Exception as e:
    st.error(f"No pude leer el archivo de Cierre: {e}")
    st.stop()

st.markdown("### 🏢 Cierre – selecciona la fila de encabezados")

header_row_cierre_1based = st.number_input(
    "Fila de encabezados en Cierre (1 = primera fila):",
    min_value=1,
    max_value=int(len(raw_cierre)),
    value=1,
    step=1,
    key="fila_header_cierre"
)
header_row_cierre = int(header_row_cierre_1based) - 1  # convertir a índice 0-based

df_cierre = build_table_from_row(raw_cierre, header_row_cierre)
df_cierre = drop_all_empty_columns(df_cierre)

st.caption("Vista previa de Cierre con los encabezados aplicados según la fila seleccionada:")
st.dataframe(df_cierre.head(20), use_container_width=True)

# ---- Balance: leer sin encabezado y seleccionar fila de header ----
try:
    raw_balance = pd.read_excel(balance_file, header=None, dtype=str)
except Exception as e:
    st.error(f"No pude leer el archivo de Balance: {e}")
    st.stop()

st.markdown("### 📊 Balance – selecciona la fila de encabezados")

header_row_balance_1based = st.number_input(
    "Fila de encabezados en Balance (1 = primera fila):",
    min_value=1,
    max_value=int(len(raw_balance)),
    value=1,
    step=1,
    key="fila_header_balance"
)
header_row_balance = int(header_row_balance_1based) - 1  # convertir a índice 0-based

df_balance = build_table_from_row(raw_balance, header_row_balance)
df_balance = drop_all_empty_columns(df_balance)

st.caption("Vista previa de Balance con los encabezados aplicados según la fila seleccionada:")
st.dataframe(df_balance.head(20), use_container_width=True)

# =========================
# 2. Configuración Cierre
# =========================
st.subheader("2. Configuración de Cierre (bloque, código y valor cobro)")

cols_cierre = list(df_cierre.columns)

# Sugerencias para Inmueble Código y Bloque
auto_codigo_cierre = find_col_fuzzy(df_cierre, ["inmueble codigo", "inmueble código", "codigo", "código"])
auto_bloque_cierre = find_col_fuzzy(df_cierre, ["inmueble bloque", "bloque", "torre"])

codigo_cierre_col = st.selectbox(
    "Columna Inmueble Código (Cierre)",
    options=cols_cierre,
    index=cols_cierre.index(auto_codigo_cierre) if auto_codigo_cierre in cols_cierre else 0
)

bloque_cierre_col = st.selectbox(
    "Columna Inmueble Bloque (Cierre)",
    options=cols_cierre,
    index=cols_cierre.index(auto_bloque_cierre) if auto_bloque_cierre in cols_cierre else 0
)

# Sugerencia para Valor Cobro
auto_valor_cobro = pick_amount_col(
    df_cierre,
    ["valor cobro", "valor a cobrar", "valor cobrado", "cobro", "cuota", "facturado", "valor"]
)

valor_cobro_col = st.selectbox(
    "Columna Valor Cobro (Cierre)",
    options=cols_cierre,
    index=cols_cierre.index(auto_valor_cobro) if auto_valor_cobro in cols_cierre else 0
)

# Construcción de clave de apartamento en Cierre: piso-num a partir de Bloque + Código
bloque_raw = df_cierre[bloque_cierre_col].astype(str).str.extract(r"(\d+)")[0]
codigo_raw = df_cierre[codigo_cierre_col].astype(str).str.extract(r"(\d{3,5})")[0]

piso = bloque_raw.str.lstrip("0")
piso = piso.replace("", np.nan)

df_cierre["_apto_key_raw"] = np.where(
    piso.notna() & codigo_raw.notna(),
    piso + "-" + codigo_raw,
    np.nan
)
df_cierre["_apto_key"] = df_cierre["_apto_key_raw"].apply(normalize_apto_key)

# Montos Cierre
df_cierre["_valor_cobro_num"] = to_amount(df_cierre[valor_cobro_col])

# Filtramos rows con clave válida
df_cierre_valid = df_cierre[df_cierre["_apto_key"].apply(is_valid_apto_format)].copy()

# Agregado Cierre por apartamento
g_cierre = (
    df_cierre_valid
    .dropna(subset=["_apto_key"])
    .groupby("_apto_key", as_index=False)
    .agg(
        valor_cobro_sum=("_valor_cobro_num", "sum"),
        conteo_registros=(codigo_cierre_col, "count")
    )
)

# =========================
# 3. Configuración Balance
# =========================
st.subheader("3. Configuración de Balance (clave apto, Nuevo Saldo y Cuenta)")

cols_balance = list(df_balance.columns)

# Columna de clave apto en Balance
auto_apto_balance = pick_apto_col_by_pattern(df_balance)
if auto_apto_balance is None:
    auto_apto_balance = find_col_fuzzy(df_balance, ["nit", "nombre nit", "apto", "apart", "unidad", "inmueble"])

apto_balance_col = st.selectbox(
    "Columna clave de apartamento (Balance, ej. NIT con 1-101)",
    options=cols_balance,
    index=cols_balance.index(auto_apto_balance) if auto_apto_balance in cols_balance else 0
)

# Columna Nuevo Saldo
auto_nuevo_saldo = next(
    (c for c in cols_balance if normalize_text(c) == "nuevo saldo"),
    pick_amount_col(df_balance, ["nuevo saldo", "saldo nuevo", "saldo final", "saldo", "balance", "cartera", "deuda"])
)

nuevo_saldo_col = st.selectbox(
    "Columna Nuevo Saldo (Balance)",
    options=cols_balance,
    index=cols_balance.index(auto_nuevo_saldo) if auto_nuevo_saldo in cols_balance else 0
)

# Columna Cuenta (para filtrar 1345*)
auto_cuenta = find_col_fuzzy(df_balance, ["cuenta", "codigo", "código", "cod"])
cuenta_col = st.selectbox(
    "Columna Cuenta (Balance, para cuentas 1345*)",
    options=cols_balance,
    index=cols_balance.index(auto_cuenta) if auto_cuenta in cols_balance else 0
)

# Clave apartamento en Balance
df_balance["_apto_key"] = df_balance[apto_balance_col].apply(normalize_apto_key)
df_balance["_apto_key_valid"] = df_balance["_apto_key"].apply(is_valid_apto_format)

# Montos Balance
df_balance["_nuevo_saldo_num"] = to_amount(df_balance[nuevo_saldo_col])

# Flag cuentas 1345*
df_balance["_is_1345"] = (
    df_balance[cuenta_col]
    .astype(str)
    .str.strip()
    .str.startswith("1345")
)

df_balance["_nuevo_saldo_1345_num"] = np.where(
    df_balance["_is_1345"],
    df_balance["_nuevo_saldo_num"],
    0.0
)

df_balance_valid = df_balance[df_balance["_apto_key_valid"]].copy()

# Agregado Balance por apartamento (solo cuentas 1345*)
g_balance = (
    df_balance_valid
    .dropna(subset=["_apto_key"])
    .groupby("_apto_key", as_index=False)
    .agg(
        nuevo_saldo_1345_sum=("_nuevo_saldo_1345_num", "sum"),
        conteo_registros=(cuenta_col, "count")
    )
)

# =========================
# 4. Match y conciliación
# =========================
st.subheader("4. Conciliación Cierre vs Balance")

res = pd.merge(g_cierre, g_balance, on="_apto_key", how="outer")

res["valor_cobro_sum"] = res["valor_cobro_sum"].fillna(0.0)
res["nuevo_saldo_1345_sum"] = res["nuevo_saldo_1345_sum"].fillna(0.0)

res["diferencia"] = res["valor_cobro_sum"] - res["nuevo_saldo_1345_sum"]

tolerance = 0.01
conciliacion = res[res["diferencia"].abs() > tolerance].sort_values("_apto_key")

# =========================
# 5. Resultados y descarga
# =========================
st.markdown("### Resultados")

m1, m2, m3, m4 = st.columns(4)
with m1:
    st.metric("Aptos en Cierre", int(g_cierre.shape[0]))
with m2:
    st.metric("Aptos en Balance (1345*)", int(g_balance.shape[0]))
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
    st.dataframe(g_cierre.sort_values("_apto_key").reset_index(drop=True), use_container_width=True)

with tabs[3]:
    st.dataframe(g_balance.sort_values("_apto_key").reset_index(drop=True), use_container_width=True)

def build_output_excel() -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as w:
        g_cierre.rename(columns={"_apto_key": "apto_key"}).to_excel(w, "agregado_cierre", index=False)
        g_balance.rename(columns={"_apto_key": "apto_key"}).to_excel(w, "agregado_balance", index=False)
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

with st.expander("Diagnóstico (configuración usada)"):
    st.json({
        "cierre": {
            "fila_encabezado": int(header_row_cierre_1based),
            "codigo_cierre_col": codigo_cierre_col,
            "bloque_cierre_col": bloque_cierre_col,
            "valor_cobro_col": valor_cobro_col,
        },
        "balance": {
            "fila_encabezado": int(header_row_balance_1based),
            "apto_balance_col": apto_balance_col,
            "nuevo_saldo_col": nuevo_saldo_col,
            "cuenta_col": cuenta_col,
        }
    })

st.caption("Cierre vs Balance por apto, sumando solo cuentas 1345* como Ingresos por Cobrar. Ahora la vista previa se actualiza cada vez que cambias la fila de encabezado 🧠📊.")
