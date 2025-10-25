# app_conciliacion.py
# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import unicodedata
from typing import Optional, List, Tuple

# =========================
# Config & Título
# =========================
st.set_page_config(page_title="Conciliación de Cartera", layout="wide")
st.title("Conciliación de Cartera: Cierre vs Balance por Terceros (Override específico)")
st.caption("Este script usa override fijo: Cartera header=8, Balance header=3. Match por 'piso-num' (ej. 1-9803).")

# =========================
# Utilidades
# =========================
def normalize_text(s: str) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    s = str(s)
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
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

def row_keyword_score(row_vals, must_have_any, must_have_optional):
    row_vals = [normalize_text(v) for v in row_vals]
    score_any = sum(any(kw in cell for cell in row_vals) for kw in must_have_any)
    score_opt = sum(any(kw in cell for cell in row_vals) for kw in must_have_optional)
    return score_any * 10 + score_opt, score_any

def build_table(df_raw: pd.DataFrame, header_row_idx: int,
                must_have_any: List[str], must_have_optional: List[str]) -> Tuple[pd.DataFrame, int]:
    """Construye tabla usando 'header_row_idx' ya definido (override)."""
    headers = df_raw.iloc[header_row_idx].astype(str).tolist()
    headers = [h if normalize_text(h) not in ("", "unnamed: 0", "nan") else f"col_{i}" for i, h in enumerate(headers)]
    headers = make_unique(headers)

    data = df_raw.iloc[header_row_idx+1:].copy()
    data.columns = headers
    data = data.dropna(how="all")

    # Limpiar espacios SOLO en object
    for c in list(data.select_dtypes(include=["object"]).columns):
        data[c] = data[c].astype(str).str.strip()

    return data, header_row_idx

def drop_all_empty_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Elimina columnas completamente vacías (NaN o strings vacíos)."""
    tmp = df.copy()
    for c in tmp.columns:
        if tmp[c].dtype == object:
            tmp[c] = tmp[c].replace("", np.nan)
    mask_nonempty = tmp.notna().any(axis=0)
    return df.loc[:, mask_nonempty.values]

def find_col_fuzzy(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    cols = list(df.columns)
    ncols = [normalize_text(c) for c in cols]
    best, score_best = None, -1
    for i, nc in enumerate(ncols):
        score = sum(kw in nc for kw in candidates)
        if score > score_best:
            best, score_best = cols[i], score
    return best

def to_amount(series: pd.Series) -> pd.Series:
    s = series.fillna("").astype(str)
    s = s.replace({r'[^0-9\-,\.]': ''}, regex=True)
    # ES: miles con punto, decimales con coma
    s = s.str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
    return pd.to_numeric(s, errors='coerce')

# --- Clave piso-num ---
APT_SEP_REGEX = r"[-_/\.]"  # separadores aceptados: -, _, /, .
def normalize_apto_key(s: str) -> Optional[str]:
    """
    Extrae 'piso-num' con separadores flexibles y lo normaliza sin ceros a la izquierda.
    Ej.: '1 - 9803' -> '1-9803', '2_9901' -> '2-9901'
    """
    if pd.isna(s):
        return None
    text = str(s).strip()
    m = re.search(rf"(\d+)\s*{APT_SEP_REGEX}\s*(\d+)", text)
    if not m:
        return None
    piso = int(m.group(1))
    num  = int(m.group(2))
    return f"{piso}-{num}"

def extract_unit_only(s: str) -> Optional[int]:
    """Fallback: devuelve solo el número de unidad (parte derecha) o la secuencia numérica más larga."""
    if pd.isna(s):
        return None
    text = str(s)
    m = re.search(rf"(\d+)\s*{APT_SEP_REGEX}\s*(\d+)", text)
    if m:
        try:
            return int(m.group(2))
        except:
            pass
    nums = re.findall(r"\d+", text)
    if not nums:
        return None
    return int(max(nums, key=len))

# =========================
# Sidebar (parámetros)
# =========================
st.sidebar.header("Parámetros")
tolerance = st.sidebar.number_input("Tolerancia (abs) para diferencia ≠ 0", min_value=0.0, value=0.01, step=0.01, format="%.2f")

st.sidebar.markdown("---")
st.sidebar.subheader("Clave de conciliación")
key_mode = st.sidebar.radio("Modo de clave para conciliar", ["Piso-Numero (ej. 1-9803)", "Solo Numero"], index=0)

st.sidebar.markdown("---")
st.sidebar.subheader("Limpieza visual")
clean_cierre = st.sidebar.checkbox("Eliminar columnas vacías en Cierre", value=True)
clean_balance = st.sidebar.checkbox("Eliminar columnas vacías en Balance", value=True)

# =========================
# Carga de archivo
# =========================
uploaded = st.file_uploader("Sube tu Excel (dos hojas: Cierre y Balance)", type=["xlsx"])
if not uploaded:
    st.info("Sube un archivo para comenzar. La app espera dos hojas: **Cierre** (apartamento, valor cobro) y **Balance** (NIT/Nombre, nuevo saldo).")
    st.stop()

# Leer libro y ofrecer selección de hojas
try:
    xls = pd.ExcelFile(uploaded)
    sheet_names = xls.sheet_names
    if len(sheet_names) < 2:
        st.error("Se requieren al menos 2 hojas.")
        st.stop()
    st.success(f"Hojas detectadas: {', '.join(sheet_names)}")
except Exception as e:
    st.error(f"No pude leer el Excel: {e}")
    st.stop()

col_sel1, col_sel2 = st.columns(2)
with col_sel1:
    sheet_cierre = st.selectbox("Hoja de Cierre (apartamento, valor cobro)", options=sheet_names, index=0)
with col_sel2:
    sheet_balance = st.selectbox("Hoja de Balance (NIT/Nombre, nuevo saldo)", options=sheet_names, index=min(1, len(sheet_names)-1))

# Lectura cruda (sin header)
raw1 = pd.read_excel(uploaded, sheet_name=sheet_cierre, header=None, dtype=str)
raw2 = pd.read_excel(uploaded, sheet_name=sheet_balance, header=None, dtype=str)

# =========================
# OVERRIDE MANUAL (caso específico)
# Cartera header en fila 8 (index 7)
# Balance header en fila 3 (index 2)
# =========================
HDR_CIERRE_IDX = 7
HDR_BALANCE_IDX = 2

df1, hdr1_used = build_table(
    raw1, HDR_CIERRE_IDX,
    must_have_any=["apartamento", "apto", "nro", "numero"],
    must_have_optional=["valor", "cobro", "cuota", "facturado"]
)
df2, hdr2_used = build_table(
    raw2, HDR_BALANCE_IDX,
    must_have_any=["nit", "tercero", "identificacion", "documento", "nombre"],
    must_have_optional=["saldo", "cartera", "balance", "credit", "debit"]
)

# Limpieza visual (sin promover encabezados, porque ya fijamos las filas)
if clean_cierre:
    df1 = drop_all_empty_columns(df1)
if clean_balance:
    df2 = drop_all_empty_columns(df2)

# =========================
# Vista previa tras override
# =========================
st.markdown("### Vista previa (override aplicado)")
st.write(f"**{sheet_balance}** (fila encabezado usada: {hdr2_used})")
st.dataframe(df2.head(12), use_container_width=True)
st.write(f"**{sheet_cierre}** (fila encabezado usada: {hdr1_used})")
st.dataframe(df1.head(12), use_container_width=True)

# =========================
# Mapeo de columnas
# =========================
apto_col_1_auto = find_col_fuzzy(df1, ["nro apartamento", "nro aptos", "no apartamento", "numero apartamento", "num apartamento", "apto", "apartamento", "inmueble", "inmueble codigo"])
valor_cobro_col_auto = find_col_fuzzy(df1, ["valor cobro", "valor a cobrar", "valor cobrado", "valor", "cobro", "cuota", "facturado"])

# En Balance, el 'piso-num' puede venir en NIT o en 'Nombre NIT' u otra columna
nuevo_saldo_col_auto = find_col_fuzzy(df2, ["nuevo saldo", "saldo nuevo", "saldo", "balance", "deuda", "cartera", "saldo final"])
apto_balance_auto = find_col_fuzzy(df2, ["apto", "apart", "unidad", "inmueble", "nombre nit", "nombre", "detalle", "nit"])

st.markdown("### Mapeo de columnas")
c1, c2 = st.columns(2)
with c1:
    apto_col_1 = st.selectbox(
        "Cierre: columna con 'piso-num' o número de apartamento",
        options=df1.columns.tolist(),
        index=(df1.columns.tolist().index(apto_col_1_auto) if apto_col_1_auto in df1.columns else 0),
    )
    valor_cobro_col = st.selectbox(
        "Cierre: columna 'Valor Cobro'",
        options=df1.columns.tolist(),
        index=(df1.columns.tolist().index(valor_cobro_col_auto) if valor_cobro_col_auto in df1.columns else 0),
    )
with c2:
    # Columna que verdaderamente contiene el 'piso-num' en Balance (puede ser NIT o Nombre NIT)
    apto_balance_candidates = df2.columns.tolist()
    apto_balance_col = st.selectbox(
        "Balance: columna con 'piso-num' (ej. 1-9803) — puede ser NIT o 'Nombre NIT'",
        options=apto_balance_candidates,
        index=(apto_balance_candidates.index(apto_balance_auto) if apto_balance_auto in apto_balance_candidates else 0),
    )
    nuevo_saldo_col = st.selectbox(
        "Balance: columna 'Nuevo Saldo'",
        options=df2.columns.tolist(),
        index=(df2.columns.tolist().index(nuevo_saldo_col_auto) if nuevo_saldo_col_auto in df2.columns else 0),
    )

# Validación
missing = []
for df_name, df_ref, col in [
    ("Cierre", df1, apto_col_1),
    ("Cierre", df1, valor_cobro_col),
    ("Balance", df2, apto_balance_col),
    ("Balance", df2, nuevo_saldo_col),
]:
    if col not in df_ref.columns:
        missing.append(f"{df_name}: {col}")
if missing:
    st.error("Columnas seleccionadas no existen tras la limpieza: " + ", ".join(missing))
    st.stop()

# =========================
# Conciliación
# =========================
# 1) Montos
df1["_valor_cobro_num"] = to_amount(df1[valor_cobro_col])
df2["_nuevo_saldo_num"] = to_amount(df2[nuevo_saldo_col])

# 2) Clave de apto (Cierre y Balance)
if key_mode.startswith("Piso-Numero"):
    df1["_apto_key"] = df1[apto_col_1].apply(normalize_apto_key)
    df2["_apto_key"] = df2[apto_balance_col].apply(normalize_apto_key)
else:
    df1["_apto_key"] = df1[apto_col_1].apply(extract_unit_only)
    df2["_apto_key"] = df2[apto_balance_col].apply(extract_unit_only)

# 3) Agregaciones
g1 = (
    df1.dropna(subset=["_apto_key"])
       .groupby("_apto_key", as_index=False)
       .agg(valor_cobro_sum=("_valor_cobro_num", "sum"),
            conteo_registros=(apto_col_1, "count"))
)
g2 = (
    df2.dropna(subset=["_apto_key"])
       .groupby("_apto_key", as_index=False)
       .agg(nuevo_saldo_sum=("_nuevo_saldo_num", "sum"))
)

# 4) Join y diferencia
res = pd.merge(g1, g2, on="_apto_key", how="outer")
res["valor_cobro_sum"] = res["valor_cobro_sum"].fillna(0.0)
res["nuevo_saldo_sum"] = res["nuevo_saldo_sum"].fillna(0.0)
res["diferencia"] = res["valor_cobro_sum"] - res["nuevo_saldo_sum"]

# 5) Filtrar diferencias ≠ 0 (tolerancia)
conciliacion = res[res["diferencia"].abs() > tolerance].sort_values("_apto_key")

# =========================
# Resultados y descarga
# =========================
st.markdown("### Resultados")
m1, m2, m3, m4 = st.columns(4)
with m1: st.metric("Aptos en Cierre", int(g1.shape[0]))
with m2: st.metric("Aptos en Balance", int(g2.shape[0]))
with m3: st.metric("Coincidencias (outer join)", int(res.shape[0]))
with m4: st.metric("Diferencias ≠ 0", int(conciliacion.shape[0]))

tabs = st.tabs(["Conciliación", "Match Total", "Agregado Cierre", "Agregado Balance"])
with tabs[0]: st.dataframe(conciliacion.reset_index(drop=True), use_container_width=True)
with tabs[1]: st.dataframe(res.sort_values("_apto_key").reset_index(drop=True), use_container_width=True)
with tabs[2]: st.dataframe(g1.sort_values("_apto_key").reset_index(drop=True), use_container_width=True)
with tabs[3]: st.dataframe(g2.sort_values("_apto_key").reset_index(drop=True), use_container_width=True)

def build_output_excel() -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as w:
        g1.rename(columns={"_apto_key": "apto_key"}).to_excel(w, "agregado_hoja1", index=False)
        g2.rename(columns={"_apto_key": "apto_key"}).to_excel(w, "agregado_hoja2", index=False)
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

with st.expander("Diagnóstico"):
    st.json({
        "sheet_cierre": sheet_cierre, "hdr_cierre_usado": hdr1_used,
        "sheet_balance": sheet_balance, "hdr_balance_usado": hdr2_used,
        "clean_cierre": clean_cierre, "clean_balance": clean_balance,
        "key_mode": key_mode,
        "apto_col_1": apto_col_1,
        "apto_balance_col": apto_balance_col,
        "valor_cobro_col": valor_cobro_col,
        "nuevo_saldo_col": nuevo_saldo_col,
        "tolerance": tolerance,
    })

st.caption("Override aplicado: Cartera header=8, Balance header=3. Si el mapeo luce raro, revisa las columnas seleccionadas.")
