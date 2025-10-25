# app_conciliacion.py
# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import unicodedata
from typing import Optional, Tuple, List
from pandas.api.types import is_string_dtype

st.set_page_config(page_title="Conciliación de Cartera", layout="wide")
st.title("Conciliación de Cartera: Cierre vs Balance por Terceros")
st.caption("Sube un Excel con dos hojas (Cierre y Balance). Detecto encabezados, mapeo columnas y genero la hoja 'conciliacion'.")

# ---------------------------
# Utilidades
# ---------------------------
def normalize_text(s: str) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    s = str(s)
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    return s.lower().strip()

def make_unique(names: List[str]) -> List[str]:
    """Devuelve una lista de nombres únicos preservando orden: a, a -> a, a_1, a_2..."""
    seen = {}
    out = []
    for n in names:
        n = str(n)
        if n not in seen:
            seen[n] = 0
            out.append(n)
        else:
            seen[n] += 1
            out.append(f"{n}_{seen[n]}")
    return out

def find_header_row(df_raw: pd.DataFrame, must_have_any: List[str], must_have_optional: Optional[List[str]] = None) -> Optional[int]:
    if must_have_optional is None:
        must_have_optional = []
    best_row, best_score = None, -1
    for i in range(len(df_raw)):
        row_vals = [normalize_text(v) for v in df_raw.iloc[i].tolist()]
        score_any = sum(any(kw in cell for cell in row_vals) for kw in must_have_any)
        score_opt = sum(any(kw in cell for cell in row_vals) for kw in must_have_optional)
        score = score_any * 10 + score_opt
        if score_any >= 1 and score > best_score:
            best_row, best_score = i, score
    return best_row

def build_table(df_raw: pd.DataFrame, header_row_idx: int) -> pd.DataFrame:
    # 1) Leer encabezados
    headers = df_raw.iloc[header_row_idx].astype(str).tolist()
    # 2) Rellenar vacíos/unnamed y hacer únicos
    headers = [h if normalize_text(h) not in ("", "unnamed: 0") else f"col_{i}" for i, h in enumerate(headers)]
    headers = make_unique(headers)
    # 3) Asignar datos y columnas
    data = df_raw.iloc[header_row_idx+1:].copy()
    data.columns = headers
    # 4) Drop de filas totalmente vacías
    data = data.dropna(how="all")
    # 5) Limpiar espacios SOLO en columnas tipo object (robusto, sin usar .dtype directo sobre DataFrame)
    obj_cols = list(data.select_dtypes(include=["object"]).columns)
    for c in obj_cols:
        data[c] = data[c].astype(str).str.strip()
    return data

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
    # Manejo ES: puntos de miles, coma decimal
    s = s.str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
    return pd.to_numeric(s, errors='coerce')

def extract_digits_longest(s: str) -> Optional[int]:
    if pd.isna(s):
        return None
    s = str(s)
    matches = re.findall(r'\d+', s)
    if not matches:
        return None
    longest = max(matches, key=len)
    try:
        return int(longest)
    except:
        return None

def extract_last_n_digits(s: str, n: int) -> Optional[int]:
    if pd.isna(s):
        return None
    s = "".join(re.findall(r'\d', str(s)))
    if len(s) < n:
        return None
    try:
        return int(s[-n:])
    except:
        return None

def extract_regex(s: str, pattern: str) -> Optional[int]:
    if pd.isna(s):
        return None
    m = re.search(pattern, str(s))
    if not m:
        return None
    try:
        return int(re.sub(r"\D", "", m.group(0)))
    except:
        return None

# ---------------------------
# Sidebar: parámetros
# ---------------------------
st.sidebar.header("Parámetros")
tolerance = st.sidebar.number_input("Tolerancia (absoluta) para considerar diferencia ≠ 0", min_value=0.0, value=0.01, step=0.01, format="%.2f")

extract_strategy = st.sidebar.selectbox(
    "Estrategia para derivar 'nro apartamento' desde el NIT (Hoja 2)",
    ["Secuencia numérica más larga", "Últimos N dígitos", "Regex personalizada"]
)

last_n = None
regex_pat = None
if extract_strategy == "Últimos N dígitos":
    last_n = st.sidebar.number_input("N (últimos N dígitos)", min_value=1, max_value=12, value=3, step=1)
elif extract_strategy == "Regex personalizada":
    regex_pat = st.sidebar.text_input("Patrón regex (ej. r'APTO\\s*(\\d{2,4})')", value=r"\d{2,6}")

st.sidebar.markdown("---")
st.sidebar.caption("La app detecta encabezados por palabras clave aunque no estén en la fila 1.")

# ---------------------------
# Carga de archivo
# ---------------------------
uploaded = st.file_uploader("Sube tu archivo Excel (xlsx)", type=["xlsx"])

if not uploaded:
    st.info("Sube un archivo para comenzar. La app espera dos hojas: una con **Cierre** (apartamento, valor cobro) y otra con **Balance** (NIT, nuevo saldo).")
    st.stop()

# Intento de lectura del Excel
try:
    xls = pd.ExcelFile(uploaded)
    sheet_names = xls.sheet_names
    if len(sheet_names) < 2:
        st.error("Se requieren al menos 2 hojas en el Excel.")
        st.stop()
    st.success(f"Archivo detectado. Hojas: {', '.join(sheet_names)}")
except Exception as e:
    st.error(f"No pude leer el Excel: {e}")
    st.stop()

# Selección de hojas (por si el orden no es el esperado)
col_sel1, col_sel2 = st.columns(2)
with col_sel1:
    sheet_cierre = st.selectbox("Hoja de Cierre (nro apartamento, valor cobro)", options=sheet_names, index=0)
with col_sel2:
    sheet_balance = st.selectbox("Hoja de Balance (NIT, nuevo saldo)", options=sheet_names, index=min(1, len(sheet_names)-1))

# Lectura cruda (sin header)
raw1 = pd.read_excel(uploaded, sheet_name=sheet_cierre, header=None, dtype=str)
raw2 = pd.read_excel(uploaded, sheet_name=sheet_balance, header=None, dtype=str)

# Heurísticas para ubicar encabezados
hdr1_idx = find_header_row(raw1, must_have_any=["apartamento", "apto", "nro", "numero"], must_have_optional=["valor", "cobro", "cuota", "facturado"])
hdr2_idx = find_header_row(raw2, must_have_any=["nit", "tercero", "identificacion", "documento"], must_have_optional=["saldo", "cartera", "balance"])

df1 = build_table(raw1, hdr1_idx) if hdr1_idx is not None else pd.read_excel(uploaded, sheet_name=sheet_cierre)
df2 = build_table(raw2, hdr2_idx) if hdr2_idx is not None else pd.read_excel(uploaded, sheet_name=sheet_balance)

st.markdown("### Vista previa de tablas detectadas")
st.write(f"**{sheet_cierre}** (fila encabezado detectada: {hdr1_idx if hdr1_idx is not None else 'no detectada, se asumió la primera fila'})")
st.dataframe(df1.head(10))
st.write(f"**{sheet_balance}** (fila encabezado detectada: {hdr2_idx if hdr2_idx is not None else 'no detectada, se asumió la primera fila'})")
st.dataframe(df2.head(10))

# Detección difusa de columnas
apto_col_1_auto = find_col_fuzzy(df1, ["nro apartamento", "nro apartamentos", "no apartamento", "numero apartamento", "num apartamento", "apto", "apartamento", "inmueble"])
valor_cobro_col_auto = find_col_fuzzy(df1, ["valor cobro", "valor a cobrar", "valor cobrado", "valor", "cobro", "cuota", "facturado"])

nit_col_2_auto = find_col_fuzzy(df2, ["nit", "identificacion", "id tercero", "tercero", "documento"])
nuevo_saldo_col_auto = find_col_fuzzy(df2, ["nuevo saldo", "saldo nuevo", "saldo", "balance", "deuda", "cartera"])

# Selección manual con defaults auto-detectados
st.markdown("### Mapeo de columnas")
c1, c2 = st.columns(2)
with c1:
    apto_col_1 = st.selectbox("Hoja Cierre: columna de nro apartamento", options=df1.columns.tolist(),
                               index=(df1.columns.tolist().index(apto_col_1_auto) if apto_col_1_auto in df1.columns else 0))
    valor_cobro_col = st.selectbox("Hoja Cierre: columna valor cobro", options=df1.columns.tolist(),
                                    index=(df1.columns.tolist().index(valor_cobro_col_auto) if valor_cobro_col_auto in df1.columns else 0))
with c2:
    nit_col_2 = st.selectbox("Hoja Balance: columna NIT / Tercero", options=df2.columns.tolist(),
                              index=(df2.columns.tolist().index(nit_col_2_auto) if nit_col_2_auto in df2.columns else 0))
    nuevo_saldo_col = st.selectbox("Hoja Balance: columna nuevo saldo", options=df2.columns.tolist(),
                                    index=(df2.columns.tolist().index(nuevo_saldo_col_auto) if nuevo_saldo_col_auto in df2.columns else 0))

# Validación de columnas seleccionadas
missing = []
for df_name, df_ref, col in [("Cierre", df1, apto_col_1), ("Cierre", df1, valor_cobro_col),
                             ("Balance", df2, nit_col_2), ("Balance", df2, nuevo_saldo_col)]:
    if col not in df_ref.columns:
        missing.append(f"{df_name}: {col}")
if missing:
    st.error("Columnas seleccionadas no existen tras normalización de encabezados: " + ", ".join(missing))
    st.stop()

# ---------------------------
# Transformaciones
# ---------------------------
# 1) Normalización de montos
df1["_valor_cobro_num"] = to_amount(df1[valor_cobro_col])
df2["_nuevo_saldo_num"] = to_amount(df2[nuevo_saldo_col])

# 2) Clave de apto en Hoja 1 (explicita o embebida en texto)
def parse_apto_from_df1(x):
    matches = re.findall(r'\d+', str(x))
    if not matches:
        return None
    try:
        return int(max(matches, key=len))
    except:
        return None

df1["_apto_num"] = df1[apto_col_1].apply(parse_apto_from_df1)

# 3) Clave de apto en Hoja 2 (derivada del NIT / Tercero)
if extract_strategy == "Secuencia numérica más larga":
    df2["_apto_num"] = df2[nit_col_2].apply(extract_digits_longest)
elif extract_strategy == "Últimos N dígitos":
    df2["_apto_num"] = df2[nit_col_2].apply(lambda s: extract_last_n_digits(s, last_n))
else:  # Regex personalizada
    df2["_apto_num"] = df2[nit_col_2].apply(lambda s: extract_regex(s, regex_pat))

# 4) Agregaciones y match
g1 = (
    df1.dropna(subset=["_apto_num"])
       .groupby("_apto_num", as_index=False)
       .agg(
           valor_cobro_sum=("_valor_cobro_num", "sum"),
           conteo_registros=(apto_col_1, "count")
       )
)
g2 = (
    df2.dropna(subset=["_apto_num"])
       .groupby("_apto_num", as_index=False)
       .agg(nuevo_saldo_sum=("_nuevo_saldo_num", "sum"))
)

res = pd.merge(g1, g2, on="_apto_num", how="outer")
res["valor_cobro_sum"] = res["valor_cobro_sum"].fillna(0.0)
res["nuevo_saldo_sum"] = res["nuevo_saldo_sum"].fillna(0.0)
res["diferencia"] = res["valor_cobro_sum"] - res["nuevo_saldo_sum"]

conciliacion = res[res["diferencia"].abs() > tolerance].copy()
conciliacion = conciliacion.sort_values(by="_apto_num")

# ---------------------------
# Resultados y descarga
# ---------------------------
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

tabs = st.tabs(["Conciliación", "Match Total", "Agregado Hoja 1", "Agregado Hoja 2"])
with tabs[0]:
    st.dataframe(conciliacion.reset_index(drop=True))
with tabs[1]:
    st.dataframe(res.sort_values("_apto_num").reset_index(drop=True))
with tabs[2]:
    st.dataframe(g1.sort_values("_apto_num").reset_index(drop=True))
with tabs[3]:
    st.dataframe(g2.sort_values("_apto_num").reset_index(drop=True))

# Construir archivo Excel de salida en memoria
def build_output_excel() -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        g1_export = g1.rename(columns={"_apto_num": "apto_num"})
        g2_export = g2.rename(columns={"_apto_num": "apto_num"})
        res_export = res.rename(columns={"_apto_num": "apto_num"})
        conciliacion_export = conciliacion.rename(columns={"_apto_num": "apto_num"})
        # Redondeo suave
        for df_exp in (g1_export, g2_export, res_export, conciliacion_export):
            for c in df_exp.columns:
                if df_exp[c].dtype.kind in "f":
                    df_exp[c] = df_exp[c].round(2)
        g1_export.to_excel(writer, sheet_name="agregado_hoja1", index=False)
        g2_export.to_excel(writer, sheet_name="agregado_hoja2", index=False)
        res_export.to_excel(writer, sheet_name="match_total", index=False)
        conciliacion_export.to_excel(writer, sheet_name="conciliacion", index=False)
    return output.getvalue()

st.markdown("### Descargar resultados")
excel_bytes = build_output_excel()
st.download_button(
    label="Descargar Excel (agregados, match y conciliación)",
    data=excel_bytes,
    file_name="conciliacion_cartera.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# ---------------------------
# Diagnóstico/Debug opcional
# ---------------------------
with st.expander("Diagnóstico de detecciones (opcional)"):
    st.json({
        "sheet_cierre": sheet_cierre,
        "sheet_balance": sheet_balance,
        "hdr1_idx": hdr1_idx,
        "hdr2_idx": hdr2_idx,
        "df1_apto_col_detected": apto_col_1_auto,
        "df1_valor_cobro_col_detected": valor_cobro_col_auto,
        "df2_nit_col_detected": nit_col_2_auto,
        "df2_nuevo_saldo_col_detected": nuevo_saldo_col_auto,
        "tolerance": tolerance,
        "extract_strategy": extract_strategy,
        "last_n": last_n,
        "regex_pat": regex_pat,
    })

st.caption("Si el 803 sigue en rojo, no es bug: es la realidad tocando la puerta 🧮🙂")
