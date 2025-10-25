# app_conciliacion.py
# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import unicodedata
from typing import Optional, List

st.set_page_config(page_title="Conciliación de Cartera", layout="wide")
st.title("Conciliación de Cartera: Cierre vs Balance por Terceros")
st.caption("Detección robusta de encabezados, limpieza de columnas y conciliación con hoja 'conciliacion'.")

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

def find_header_row(df_raw: pd.DataFrame, must_have_any: List[str], must_have_optional: Optional[List[str]] = None) -> Optional[int]:
    if must_have_optional is None:
        must_have_optional = []
    best_row, best_score = None, -1
    for i in range(len(df_raw)):
        score, score_any = row_keyword_score(df_raw.iloc[i].tolist(), must_have_any, must_have_optional)
        if score_any >= 1 and score > best_score:
            best_row, best_score = i, score
    return best_row

def build_table(df_raw: pd.DataFrame, header_row_idx: int,
                must_have_any: List[str], must_have_optional: List[str]) -> (pd.DataFrame, int):
    """
    Construye tabla usando 'header_row_idx'. Si la fila siguiente luce mejor (título vs. encabezado),
    promueve automáticamente header_row_idx+1.
    """
    # ¿La siguiente fila tiene mejor pinta de encabezado?
    try:
        score_here, any_here = row_keyword_score(df_raw.iloc[header_row_idx].tolist(), must_have_any, must_have_optional)
        score_next, any_next = row_keyword_score(df_raw.iloc[header_row_idx+1].tolist(), must_have_any, must_have_optional)
        if any_next >= 1 and score_next > score_here:
            header_row_idx = header_row_idx + 1
    except Exception:
        pass

    headers = df_raw.iloc[header_row_idx].astype(str).tolist()
    headers = [h if normalize_text(h) not in ("", "unnamed: 0", "nan") else f"col_{i}" for i, h in enumerate(headers)]
    headers = make_unique(headers)

    data = df_raw.iloc[header_row_idx+1:].copy()
    data.columns = headers
    data = data.dropna(how="all")

    # strip a object
    obj_cols = list(data.select_dtypes(include=["object"]).columns)
    for c in obj_cols:
        data[c] = data[c].astype(str).str.strip()

    return data, header_row_idx

def promote_first_row_as_header(df: pd.DataFrame) -> pd.DataFrame:
    """Usa la primera fila de datos como encabezado."""
    if df.empty: 
        return df
    new_headers = df.iloc[0].astype(str).tolist()
    new_headers = [h if normalize_text(h) not in ("", "nan") else f"col_{i}" for i, h in enumerate(new_headers)]
    new_headers = make_unique(new_headers)
    df2 = df.iloc[1:].copy()
    df2.columns = new_headers
    # Limpieza post-cambio
    obj_cols = list(df2.select_dtypes(include=["object"]).columns)
    for c in obj_cols:
        df2[c] = df2[c].astype(str).str.strip()
    return df2

def drop_all_empty_columns(df: pd.DataFrame) -> pd.DataFrame:
    # Quita columnas completamente vacías (NaN o strings vacíos)
    tmp = df.copy()
    # Considerar cadenas vacías como NaN temporales
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
    s = s.str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
    return pd.to_numeric(s, errors='coerce')

def extract_digits_longest(s: str) -> Optional[int]:
    if pd.isna(s): return None
    matches = re.findall(r'\d+', str(s))
    if not matches: return None
    try:
        return int(max(matches, key=len))
    except: 
        return None

def extract_last_n_digits(s: str, n: int) -> Optional[int]:
    if pd.isna(s): return None
    s = "".join(re.findall(r'\d', str(s)))
    if len(s) < n: return None
    try:
        return int(s[-n:])
    except:
        return None

def extract_regex(s: str, pattern: str) -> Optional[int]:
    if pd.isna(s): return None
    m = re.search(pattern, str(s))
    if not m: return None
    try:
        return int(re.sub(r"\D", "", m.group(0)))
    except:
        return None

# ---------------------------
# Sidebar
# ---------------------------
st.sidebar.header("Parámetros")
tolerance = st.sidebar.number_input("Tolerancia (abs) diferencia ≠ 0", min_value=0.0, value=0.01, step=0.01, format="%.2f")

extract_strategy = st.sidebar.selectbox(
    "Derivar 'nro apartamento' desde NIT (Hoja Balance)",
    ["Secuencia numérica más larga", "Últimos N dígitos", "Regex personalizada"]
)
last_n = st.sidebar.number_input("N (últimos N dígitos)", 1, 12, 3) if extract_strategy == "Últimos N dígitos" else None
regex_pat = st.sidebar.text_input("Regex (ej. \\d{2,6})", value=r"\d{2,6}") if extract_strategy == "Regex personalizada" else None

st.sidebar.markdown("---")
st.sidebar.subheader("Limpieza visual")
clean_cierre = st.sidebar.checkbox("Eliminar columnas vacías en Cierre", value=True)
clean_balance = st.sidebar.checkbox("Eliminar columnas vacías en Balance", value=True)
promote_balance = st.sidebar.checkbox("Promover 1ª fila como encabezado (Balance)", value=True)
promote_cierre = st.sidebar.checkbox("Promover 1ª fila como encabezado (Cierre) si se ve mejor", value=False)

# ---------------------------
# Carga
# ---------------------------
uploaded = st.file_uploader("Sube tu Excel (dos hojas: Cierre y Balance)", type=["xlsx"])
if not uploaded:
    st.info("Sube un archivo para comenzar.")
    st.stop()

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
    sheet_balance = st.selectbox("Hoja de Balance (NIT, nuevo saldo)", options=sheet_names, index=min(1, len(sheet_names)-1))

raw1 = pd.read_excel(uploaded, sheet_name=sheet_cierre, header=None, dtype=str)
raw2 = pd.read_excel(uploaded, sheet_name=sheet_balance, header=None, dtype=str)

# Heurística de encabezados (ahora con doble chequeo)
hdr1_idx = find_header_row(raw1, ["apartamento", "apto", "nro", "numero"], ["valor", "cobro", "cuota", "facturado"])
hdr2_idx = find_header_row(raw2, ["nit", "tercero", "identificacion", "documento"], ["saldo", "cartera", "balance"])

df1, hdr1_used = build_table(raw1, hdr1_idx if hdr1_idx is not None else 0,
                             ["apartamento", "apto", "nro", "numero"], ["valor", "cobro", "cuota", "facturado"])
df2, hdr2_used = build_table(raw2, hdr2_idx if hdr2_idx is not None else 0,
                             ["nit", "tercero", "identificacion", "documento"], ["saldo", "cartera", "balance"])

# Limpieza visual
if promote_balance:
    df2 = promote_first_row_as_header(df2)
if promote_cierre:
    # Solo promueve si hay muchas columnas 'nan' y la primera fila parece encabezado
    if sum(normalize_text(c).startswith("col_") or normalize_text(c) == "nan" for c in df1.columns) >= max(2, len(df1.columns)//3):
        df1 = promote_first_row_as_header(df1)

if clean_cierre:
    df1 = drop_all_empty_columns(df1)
if clean_balance:
    df2 = drop_all_empty_columns(df2)

st.markdown("### Vista previa (tras limpieza)")
st.write(f"**{sheet_balance}** (fila encabezado usada: {hdr2_used})")
st.dataframe(df2.head(12))
st.write(f"**{sheet_cierre}** (fila encabezado usada: {hdr1_used})")
st.dataframe(df1.head(12))

# ---------------------------
# Mapeo de columnas (más claro ahora)
# ---------------------------
apto_col_1_auto = find_col_fuzzy(df1, ["nro apartamento", "nro aptos", "no apartamento", "numero apartamento", "num apartamento", "apto", "apartamento", "inmueble"])
valor_cobro_col_auto = find_col_fuzzy(df1, ["valor cobro", "valor a cobrar", "valor cobrado", "valor", "cobro", "cuota", "facturado"])

nit_col_2_auto = find_col_fuzzy(df2, ["nit", "identificacion", "id tercero", "tercero", "documento"])
nuevo_saldo_col_auto = find_col_fuzzy(df2, ["nuevo saldo", "saldo nuevo", "saldo", "balance", "deuda", "cartera", "saldo final"])

st.markdown("### Mapeo de columnas")
c1, c2 = st.columns(2)
with c1:
    apto_col_1 = st.selectbox("Cierre: nro apartamento", options=df1.columns.tolist(),
                               index=(df1.columns.tolist().index(apto_col_1_auto) if apto_col_1_auto in df1.columns else 0))
    valor_cobro_col = st.selectbox("Cierre: valor cobro", options=df1.columns.tolist(),
                                    index=(df1.columns.tolist().index(valor_cobro_col_auto) if valor_cobro_col_auto in df1.columns else 0))
with c2:
    nit_col_2 = st.selectbox("Balance: NIT / Tercero", options=df2.columns.tolist(),
                              index=(df2.columns.tolist().index(nit_col_2_auto) if nit_col_2_auto in df2.columns else 0))
    nuevo_saldo_col = st.selectbox("Balance: nuevo saldo", options=df2.columns.tolist(),
                                    index=(df2.columns.tolist().index(nuevo_saldo_col_auto) if nuevo_saldo_col_auto in df2.columns else 0))

# Validación
missing = []
for df_name, df_ref, col in [("Cierre", df1, apto_col_1), ("Cierre", df1, valor_cobro_col),
                             ("Balance", df2, nit_col_2), ("Balance", df2, nuevo_saldo_col)]:
    if col not in df_ref.columns:
        missing.append(f"{df_name}: {col}")
if missing:
    st.error("Columnas seleccionadas no existen tras la limpieza: " + ", ".join(missing))
    st.stop()

# ---------------------------
# Conciliación
# ---------------------------
def to_amount(series: pd.Series) -> pd.Series:
    s = series.fillna("").astype(str)
    s = s.replace({r'[^0-9\-,\.]': ''}, regex=True)
    s = s.str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
    return pd.to_numeric(s, errors='coerce')

df1["_valor_cobro_num"] = to_amount(df1[valor_cobro_col])
df2["_nuevo_saldo_num"] = to_amount(df2[nuevo_saldo_col])

def parse_apto_df1(x):
    m = re.findall(r'\d+', str(x))
    if not m: return None
    try: return int(max(m, key=len))
    except: return None

df1["_apto_num"] = df1[apto_col_1].apply(parse_apto_df1)

if extract_strategy == "Secuencia numérica más larga":
    df2["_apto_num"] = df2[nit_col_2].apply(extract_digits_longest)
elif extract_strategy == "Últimos N dígitos":
    df2["_apto_num"] = df2[nit_col_2].apply(lambda s: extract_last_n_digits(s, last_n))
else:
    df2["_apto_num"] = df2[nit_col_2].apply(lambda s: extract_regex(s, regex_pat))

g1 = (df1.dropna(subset=["_apto_num"])
         .groupby("_apto_num", as_index=False)
         .agg(valor_cobro_sum=("_valor_cobro_num","sum"),
              conteo_registros=(apto_col_1,"count")))

g2 = (df2.dropna(subset=["_apto_num"])
         .groupby("_apto_num", as_index=False)
         .agg(nuevo_saldo_sum=("_nuevo_saldo_num","sum")))

res = pd.merge(g1, g2, on="_apto_num", how="outer")
res["valor_cobro_sum"] = res["valor_cobro_sum"].fillna(0.0)
res["nuevo_saldo_sum"] = res["nuevo_saldo_sum"].fillna(0.0)
res["diferencia"] = res["valor_cobro_sum"] - res["nuevo_saldo_sum"]

conciliacion = res[res["diferencia"].abs() > tolerance].sort_values("_apto_num")

st.markdown("### Resultados")
m1, m2, m3, m4 = st.columns(4)
with m1: st.metric("Aptos en Cierre", int(g1.shape[0]))
with m2: st.metric("Aptos en Balance", int(g2.shape[0]))
with m3: st.metric("Coincidencias (outer)", int(res.shape[0]))
with m4: st.metric("Diferencias ≠ 0", int(conciliacion.shape[0]))

tabs = st.tabs(["Conciliación", "Match Total", "Agregado Cierre", "Agregado Balance"])
with tabs[0]: st.dataframe(conciliacion.reset_index(drop=True))
with tabs[1]: st.dataframe(res.sort_values("_apto_num").reset_index(drop=True))
with tabs[2]: st.dataframe(g1.sort_values("_apto_num").reset_index(drop=True))
with tabs[3]: st.dataframe(g2.sort_values("_apto_num").reset_index(drop=True))

def build_output_excel() -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as w:
        g1.rename(columns={"_apto_num":"apto_num"}).to_excel(w, "agregado_hoja1", index=False)
        g2.rename(columns={"_apto_num":"apto_num"}).to_excel(w, "agregado_hoja2", index=False)
        res.rename(columns={"_apto_num":"apto_num"}).to_excel(w, "match_total", index=False)
        conciliacion.rename(columns={"_apto_num":"apto_num"}).to_excel(w, "conciliacion", index=False)
    return output.getvalue()

st.markdown("### Descargar resultados")
st.download_button(
    "Descargar Excel (agregados, match y conciliación)",
    data=build_output_excel(),
    file_name="conciliacion_cartera.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

with st.expander("Diagnóstico de detecciones"):
    st.json({
        "sheet_balance": sheet_balance, "hdr_balance_usado": hdr2_used,
        "sheet_cierre": sheet_cierre, "hdr_cierre_usado": hdr1_used,
        "clean_cierre": clean_cierre, "clean_balance": clean_balance,
        "promote_balance": promote_balance, "promote_cierre": promote_cierre
    })
st.caption("Si antes veías 'nan' en encabezados o una columna vacía, activa los toggles de limpieza/promoción. Tu vista debería quedar limpia y fácil de mapear. 🧽✨")
