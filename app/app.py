# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import unicodedata
from typing import Optional, List

# =========================
# Config
# =========================
st.set_page_config(
    page_title="Conciliación de Cartera",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# =========================
# CSS — estilo Palantir
# =========================
st.markdown("""
<style>
/* ── Fuente y fondo ── */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

.stApp {
    background-color: #0d1117;
    color: #e6edf3;
}

/* ── Header principal ── */
.app-header {
    background: linear-gradient(135deg, #161b22 0%, #0d1117 100%);
    border-bottom: 1px solid #21262d;
    padding: 2rem 2.5rem 1.5rem;
    margin: -1rem -1rem 2rem -1rem;
}
.app-title {
    font-size: 1.6rem;
    font-weight: 700;
    color: #f0f6fc;
    letter-spacing: -0.02em;
    margin: 0 0 0.25rem 0;
}
.app-subtitle {
    font-size: 0.85rem;
    color: #8b949e;
    font-weight: 400;
    margin: 0;
}

/* ── Tarjetas de sección ── */
.section-card {
    background: #161b22;
    border: 1px solid #21262d;
    border-radius: 8px;
    padding: 1.5rem;
    margin-bottom: 1.25rem;
}
.section-title {
    font-size: 0.7rem;
    font-weight: 600;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    color: #58a6ff;
    margin: 0 0 1rem 0;
    display: flex;
    align-items: center;
    gap: 0.5rem;
}
.step-badge {
    background: #1f6feb;
    color: #fff;
    border-radius: 50%;
    width: 20px;
    height: 20px;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    font-size: 0.65rem;
    font-weight: 700;
    flex-shrink: 0;
}

/* ── Métricas customizadas ── */
.metric-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
    gap: 1px;
    background: #21262d;
    border: 1px solid #21262d;
    border-radius: 8px;
    overflow: hidden;
    margin-bottom: 1.25rem;
}
.metric-cell {
    background: #161b22;
    padding: 1rem 1.25rem;
    text-align: center;
}
.metric-cell.danger { background: #1a0f0f; }
.metric-cell.warning { background: #141008; }
.metric-cell.success { background: #0a1a0a; }
.metric-label {
    font-size: 0.7rem;
    font-weight: 500;
    color: #8b949e;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    margin-bottom: 0.4rem;
}
.metric-value {
    font-size: 1.8rem;
    font-weight: 700;
    color: #f0f6fc;
    line-height: 1;
}
.metric-value.danger { color: #f85149; }
.metric-value.warning { color: #d29922; }
.metric-value.success { color: #3fb950; }

/* ── Badges de estado ── */
.badge {
    display: inline-block;
    padding: 0.2rem 0.6rem;
    border-radius: 20px;
    font-size: 0.72rem;
    font-weight: 600;
}
.badge-ok    { background: #0d2d0d; color: #3fb950; border: 1px solid #238636; }
.badge-warn  { background: #2d2200; color: #d29922; border: 1px solid #9e6a03; }
.badge-error { background: #2d0d0d; color: #f85149; border: 1px solid #da3633; }
.badge-info  { background: #0d1f3c; color: #58a6ff; border: 1px solid #1f6feb; }

/* ── Alertas inline ── */
.alert {
    padding: 0.75rem 1rem;
    border-radius: 6px;
    font-size: 0.82rem;
    margin: 0.75rem 0;
    display: flex;
    align-items: flex-start;
    gap: 0.6rem;
}
.alert-success { background: #0a1a0a; border-left: 3px solid #3fb950; color: #3fb950; }
.alert-warning { background: #1a1500; border-left: 3px solid #d29922; color: #d29922; }
.alert-error   { background: #1a0a0a; border-left: 3px solid #f85149; color: #f85149; }
.alert-info    { background: #0d1f3c; border-left: 3px solid #58a6ff; color: #58a6ff; }

/* ── Divider ── */
.hdivider {
    height: 1px;
    background: #21262d;
    margin: 1.5rem 0;
}

/* ── Streamlit overrides ── */
.stFileUploader > div {
    background: #161b22 !important;
    border: 1px dashed #30363d !important;
    border-radius: 8px !important;
}
.stFileUploader > div:hover {
    border-color: #58a6ff !important;
}
div[data-testid="stFileUploadDropzone"] {
    background: #161b22 !important;
}

.stNumberInput > div > div > input,
.stSelectbox > div > div > div {
    background-color: #161b22 !important;
    border: 1px solid #30363d !important;
    color: #e6edf3 !important;
    border-radius: 6px !important;
}
.stSelectbox > div > div > div:hover {
    border-color: #58a6ff !important;
}

.stTabs [data-baseweb="tab-list"] {
    background: #161b22;
    border-bottom: 1px solid #21262d;
    gap: 0;
}
.stTabs [data-baseweb="tab"] {
    background: transparent;
    color: #8b949e;
    font-size: 0.8rem;
    font-weight: 500;
    padding: 0.6rem 1rem;
    border-bottom: 2px solid transparent;
}
.stTabs [aria-selected="true"] {
    color: #58a6ff !important;
    border-bottom: 2px solid #58a6ff !important;
    background: transparent !important;
}

.stDownloadButton > button {
    background: #1f6feb !important;
    color: #fff !important;
    border: none !important;
    border-radius: 6px !important;
    font-weight: 600 !important;
    font-size: 0.85rem !important;
    padding: 0.6rem 1.5rem !important;
    transition: background 0.15s !important;
}
.stDownloadButton > button:hover {
    background: #388bfd !important;
}

.stButton > button {
    background: #21262d !important;
    color: #e6edf3 !important;
    border: 1px solid #30363d !important;
    border-radius: 6px !important;
    font-size: 0.82rem !important;
}

/* Dataframe */
.stDataFrame { border-radius: 6px; overflow: hidden; }
[data-testid="stDataFrameResizable"] {
    border: 1px solid #21262d !important;
    border-radius: 6px !important;
}

/* Labels */
.stSelectbox label, .stNumberInput label, .stFileUploader label {
    color: #8b949e !important;
    font-size: 0.78rem !important;
    font-weight: 500 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.06em !important;
}

/* Subheader override */
h2, h3 { color: #f0f6fc !important; }
h4 { color: #8b949e !important; font-weight: 500 !important; font-size: 0.85rem !important; }

/* Info/warning/error boxes */
.stAlert {
    background: #161b22 !important;
    border-radius: 6px !important;
}

/* Sidebar */
section[data-testid="stSidebar"] {
    background: #161b22;
    border-right: 1px solid #21262d;
}

/* Caption */
.stCaption { color: #6e7681 !important; font-size: 0.75rem !important; }

/* Progress bar de pasos */
.stepper {
    display: flex;
    align-items: center;
    gap: 0;
    margin-bottom: 2rem;
    background: #161b22;
    border: 1px solid #21262d;
    border-radius: 8px;
    overflow: hidden;
}
.step-item {
    flex: 1;
    padding: 0.65rem 0.5rem;
    text-align: center;
    font-size: 0.7rem;
    font-weight: 600;
    color: #6e7681;
    background: #161b22;
    border-right: 1px solid #21262d;
    letter-spacing: 0.04em;
}
.step-item:last-child { border-right: none; }
.step-item.active {
    background: #0d1f3c;
    color: #58a6ff;
}
.step-item.done {
    background: #0a1a0a;
    color: #3fb950;
}
</style>
""", unsafe_allow_html=True)

# =========================
# Header
# =========================
st.markdown("""
<div class="app-header">
    <p class="app-title">Conciliación de Cartera</p>
    <p class="app-subtitle">Cierre vs Balance — cuentas 1345* &nbsp;·&nbsp; Frutto Foods</p>
</div>
""", unsafe_allow_html=True)

# =========================
# Utilidades
# =========================
def normalize_text(s: str) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    s = str(s)
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    return s.lower().strip()

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
    s = s.str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
    return pd.to_numeric(s, errors='coerce')

def build_table_from_row(df_raw: pd.DataFrame, header_row_idx: int) -> pd.DataFrame:
    headers = df_raw.iloc[header_row_idx].astype(str).tolist()
    headers_norm = []
    for i, h in enumerate(headers):
        h_norm = normalize_text(h)
        if h_norm in ("", "unnamed: 0", "nan"):
            headers_norm.append(f"col_{i}")
        else:
            headers_norm.append(h)
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
PI_SO_NUM_FULL = re.compile(rf"^\s*(\d{{1,2}})\s*{APT_SEP_REGEX}\s*(\d{{3,5}})\s*$")

def normalize_apto_key(s: str) -> Optional[str]:
    if pd.isna(s):
        return None
    text = str(s).strip()
    m = PI_SO_NUM_FULL.match(text)
    if not m:
        return None
    piso = int(m.group(1))
    num  = int(m.group(2))
    return f"{piso}-{num}"

def is_valid_apto_format(s: str) -> bool:
    if pd.isna(s):
        return False
    return bool(PI_SO_NUM_FULL.match(str(s).strip()))

def find_col_fuzzy(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    cols = list(df.columns)
    ncols = [normalize_text(c) for c in cols]
    best, score_best = None, 0
    for i, nc in enumerate(ncols):
        score = sum(kw in nc for kw in candidates)
        if score > score_best:
            best, score_best = cols[i], score
    return best  # None si score=0 en todas

def pick_apto_col_by_pattern(df: pd.DataFrame) -> Optional[str]:
    best_col, best_hits = None, -1
    sample_n = min(300, len(df))
    for col in df.columns:
        s = df[col].dropna().astype(str).head(sample_n)
        hits = s.apply(lambda x: 1 if is_valid_apto_format(x) else 0).sum()
        if hits > best_hits:
            best_hits = hits
            best_col = col
    return best_col if best_hits > 0 else None

def pick_amount_col(df: pd.DataFrame, prefer_keywords: List[str]) -> Optional[str]:
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

def fmt_cop(val) -> str:
    """Formatea número como moneda colombiana."""
    try:
        v = float(val)
        if pd.isna(v):
            return "—"
        sign = "-" if v < 0 else ""
        return f"{sign}$ {abs(v):,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(val)

def clean_display(df: pd.DataFrame) -> pd.DataFrame:
    """Elimina columnas internas (_*) y renombra las conocidas para mostrar."""
    rename_map = {
        "_apto_key": "Apartamento",
        "valor_cobro_sum": "Valor Cobro (Cierre)",
        "nuevo_saldo_1345_sum": "Saldo 1345 (Balance)",
        "diferencia": "Diferencia",
        "conteo_registros_x": "Registros Cierre",
        "conteo_registros_y": "Registros Balance",
        "conteo_registros": "Registros",
    }
    cols_to_drop = [c for c in df.columns if c.startswith("_")]
    result = df.drop(columns=cols_to_drop, errors="ignore")
    result = result.rename(columns={k: v for k, v in rename_map.items() if k in result.columns})
    return result

@st.cache_data(show_spinner=False)
def read_excel_cached(file_bytes: bytes, filename: str) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(file_bytes), header=None, dtype=str)

def section(number: int, title: str, icon: str = "") -> None:
    st.markdown(f"""
    <div class="section-title">
        <span class="step-badge">{number}</span>
        {icon}&nbsp;{title}
    </div>
    """, unsafe_allow_html=True)

def alert(msg: str, kind: str = "info") -> None:
    icons = {"success": "✓", "warning": "⚠", "error": "✕", "info": "ℹ"}
    st.markdown(f'<div class="alert alert-{kind}"><span>{icons.get(kind,"ℹ")}</span><span>{msg}</span></div>',
                unsafe_allow_html=True)

# =========================
# PASO 1 — Carga de archivos
# =========================
st.markdown('<div class="section-card">', unsafe_allow_html=True)
section(1, "CARGAR ARCHIVOS", "📂")

col_a, col_b = st.columns(2)
with col_a:
    cierre_file = st.file_uploader("Archivo de Cierre (.xlsx)", type=["xlsx"], key="cierre",
                                   help="Reporte de Cierre exportado desde el sistema de cartera")
with col_b:
    balance_file = st.file_uploader("Archivo de Balance (.xlsx)", type=["xlsx"], key="balance",
                                    help="Balance de comprobación con cuentas 1345*")

st.markdown('</div>', unsafe_allow_html=True)

if not cierre_file or not balance_file:
    if cierre_file and not balance_file:
        alert("Cierre cargado. Falta el archivo de Balance.", "warning")
    elif balance_file and not cierre_file:
        alert("Balance cargado. Falta el archivo de Cierre.", "warning")
    else:
        alert("Sube los dos archivos para comenzar la conciliación.", "info")
    st.stop()

# Leer archivos con caché
try:
    raw_cierre = read_excel_cached(cierre_file.getvalue(), cierre_file.name)
except Exception as e:
    alert(f"No se pudo leer el archivo de Cierre: {e}", "error")
    st.stop()

try:
    raw_balance = read_excel_cached(balance_file.getvalue(), balance_file.name)
except Exception as e:
    alert(f"No se pudo leer el archivo de Balance: {e}", "error")
    st.stop()

alert(f"Archivos cargados — Cierre: {len(raw_cierre)} filas · Balance: {len(raw_balance)} filas", "success")

# =========================
# PASO 2 — Encabezados
# =========================
st.markdown('<div class="section-card">', unsafe_allow_html=True)
section(2, "INDICAR FILA DE ENCABEZADOS", "🔍")
st.caption("Cada archivo puede tener filas de título o metadatos arriba de los datos reales. "
           "Mira la vista previa y ajusta el número hasta que la primera fila sean los nombres de columna.")

tab_enc_c, tab_enc_b = st.tabs(["  Cierre  ", "  Balance  "])

with tab_enc_c:
    st.markdown("**Primeras filas del archivo (tal como están en Excel):**")
    st.dataframe(raw_cierre.head(8).rename(columns=lambda x: f"col {x}"),
                 use_container_width=True, height=220)
    header_row_cierre_1based = st.number_input(
        "¿En qué fila están los NOMBRES de columna? (cuenta desde 1)",
        min_value=1, max_value=int(len(raw_cierre)), value=1, step=1,
        key="fila_header_cierre"
    )
    header_row_cierre = int(header_row_cierre_1based) - 1
    df_cierre = build_table_from_row(raw_cierre, header_row_cierre)
    df_cierre = drop_all_empty_columns(df_cierre)
    st.markdown("**Vista previa con encabezados aplicados:**")
    st.dataframe(df_cierre.head(6), use_container_width=True, height=200)

with tab_enc_b:
    st.markdown("**Primeras filas del archivo (tal como están en Excel):**")
    st.dataframe(raw_balance.head(8).rename(columns=lambda x: f"col {x}"),
                 use_container_width=True, height=220)
    header_row_balance_1based = st.number_input(
        "¿En qué fila están los NOMBRES de columna? (cuenta desde 1)",
        min_value=1, max_value=int(len(raw_balance)), value=1, step=1,
        key="fila_header_balance"
    )
    header_row_balance = int(header_row_balance_1based) - 1
    df_balance = build_table_from_row(raw_balance, header_row_balance)
    df_balance = drop_all_empty_columns(df_balance)
    st.markdown("**Vista previa con encabezados aplicados:**")
    st.dataframe(df_balance.head(6), use_container_width=True, height=200)

st.markdown('</div>', unsafe_allow_html=True)

# =========================
# PASO 3 — Columnas Cierre
# =========================
st.markdown('<div class="section-card">', unsafe_allow_html=True)
section(3, "COLUMNAS DEL CIERRE", "📋")
st.caption("La app detecta automáticamente las columnas. Corrígelas si la sugerencia no es exacta.")

cols_cierre = list(df_cierre.columns)
auto_codigo_cierre = find_col_fuzzy(df_cierre, ["inmueble codigo", "inmueble código", "codigo", "código"]) or cols_cierre[0]
auto_bloque_cierre = find_col_fuzzy(df_cierre, ["inmueble bloque", "bloque", "torre"]) or cols_cierre[0]
auto_valor_cobro   = pick_amount_col(df_cierre, ["valor cobro", "valor a cobrar", "cobro", "cuota", "facturado", "valor"]) or cols_cierre[0]

c1, c2, c3 = st.columns(3)
with c1:
    codigo_cierre_col = st.selectbox("Número del apartamento",  options=cols_cierre,
        index=cols_cierre.index(auto_codigo_cierre) if auto_codigo_cierre in cols_cierre else 0,
        help="Columna con el código numérico del apartamento (ej: 9801, 203)")
with c2:
    bloque_cierre_col = st.selectbox("Bloque / Torre",  options=cols_cierre,
        index=cols_cierre.index(auto_bloque_cierre) if auto_bloque_cierre in cols_cierre else 0,
        help="Columna con el número de bloque o torre (ej: 1, 2, 3)")
with c3:
    valor_cobro_col = st.selectbox("Valor Cobro",  options=cols_cierre,
        index=cols_cierre.index(auto_valor_cobro) if auto_valor_cobro in cols_cierre else 0,
        help="Columna con el monto cobrado o facturado")

# Construir clave y agregar
bloque_raw = df_cierre[bloque_cierre_col].astype(str).str.extract(r"(\d+)")[0]
codigo_raw = df_cierre[codigo_cierre_col].astype(str).str.extract(r"(\d{3,5})")[0]
piso = bloque_raw.str.lstrip("0").replace("", np.nan)

df_cierre["_apto_key_raw"] = np.where(
    piso.notna() & codigo_raw.notna(), piso + "-" + codigo_raw, np.nan
)
df_cierre["_apto_key"] = df_cierre["_apto_key_raw"].apply(normalize_apto_key)
df_cierre["_valor_cobro_num"] = to_amount(df_cierre[valor_cobro_col])
df_cierre_valid = df_cierre[df_cierre["_apto_key"].apply(is_valid_apto_format)].copy()

g_cierre = (
    df_cierre_valid.dropna(subset=["_apto_key"])
    .groupby("_apto_key", as_index=False)
    .agg(valor_cobro_sum=("_valor_cobro_num", "sum"),
         conteo_registros=(codigo_cierre_col, "count"))
)

n_aptos_cierre = len(g_cierre)
if n_aptos_cierre == 0:
    alert("No se detectaron apartamentos válidos con las columnas seleccionadas. "
          "Revisa las columnas Bloque y Número de apartamento.", "error")
else:
    alert(f"{n_aptos_cierre} apartamentos identificados en el Cierre.", "success")

st.markdown('</div>', unsafe_allow_html=True)

# =========================
# PASO 4 — Columnas Balance
# =========================
st.markdown('<div class="section-card">', unsafe_allow_html=True)
section(4, "COLUMNAS DEL BALANCE", "📊")
st.caption("Solo se toman las filas cuya cuenta empiece por 1345.")

cols_balance = list(df_balance.columns)
auto_apto_balance  = pick_apto_col_by_pattern(df_balance) or \
                     find_col_fuzzy(df_balance, ["apto", "apart", "unidad", "inmueble", "nit", "nombre nit"]) or \
                     cols_balance[0]
auto_nuevo_saldo   = next((c for c in cols_balance if normalize_text(c) == "nuevo saldo"), None) or \
                     pick_amount_col(df_balance, ["nuevo saldo", "saldo nuevo", "saldo final", "saldo", "balance", "cartera"]) or \
                     cols_balance[0]
auto_cuenta        = find_col_fuzzy(df_balance, ["cuenta", "codigo", "código", "cod"]) or cols_balance[0]

b1, b2, b3 = st.columns(3)
with b1:
    apto_balance_col = st.selectbox("Clave de apartamento",  options=cols_balance,
        index=cols_balance.index(auto_apto_balance) if auto_apto_balance in cols_balance else 0,
        help="Columna con la clave del apartamento en formato 1-101, 2-203, etc.")
with b2:
    nuevo_saldo_col = st.selectbox("Nuevo Saldo",  options=cols_balance,
        index=cols_balance.index(auto_nuevo_saldo) if auto_nuevo_saldo in cols_balance else 0,
        help="Columna con el saldo contable del apartamento")
with b3:
    cuenta_col = st.selectbox("Cuenta contable",  options=cols_balance,
        index=cols_balance.index(auto_cuenta) if auto_cuenta in cols_balance else 0,
        help="Columna de código contable — se filtran las que empiezan con 1345")

# Procesar balance
df_balance["_apto_key"]       = df_balance[apto_balance_col].apply(normalize_apto_key)
df_balance["_apto_key_valid"] = df_balance["_apto_key"].apply(is_valid_apto_format)
df_balance["_nuevo_saldo_num"] = to_amount(df_balance[nuevo_saldo_col])
df_balance["_is_1345"] = df_balance[cuenta_col].astype(str).str.strip().str.startswith("1345")
df_balance["_nuevo_saldo_1345_num"] = np.where(df_balance["_is_1345"], df_balance["_nuevo_saldo_num"], 0.0)
df_balance_valid = df_balance[df_balance["_apto_key_valid"]].copy()

g_balance = (
    df_balance_valid.dropna(subset=["_apto_key"])
    .groupby("_apto_key", as_index=False)
    .agg(nuevo_saldo_1345_sum=("_nuevo_saldo_1345_num", "sum"),
         conteo_registros=(cuenta_col, "count"))
)

n_aptos_balance = len(g_balance)
n_filas_1345    = df_balance["_is_1345"].sum()
if n_aptos_balance == 0:
    alert("No se detectaron apartamentos válidos en el Balance. "
          "Revisa la columna de clave de apartamento.", "error")
else:
    alert(f"{n_aptos_balance} apartamentos identificados en el Balance · {n_filas_1345} filas con cuentas 1345*", "success")

st.markdown('</div>', unsafe_allow_html=True)

# Guard: si alguno tiene 0 aptos no tiene sentido continuar
if n_aptos_cierre == 0 or n_aptos_balance == 0:
    alert("Corrige las columnas en los pasos anteriores para ver la conciliación.", "warning")
    st.stop()

# =========================
# PASO 5 — Conciliación
# =========================
res = pd.merge(g_cierre, g_balance, on="_apto_key", how="outer")
res["valor_cobro_sum"]      = res["valor_cobro_sum"].fillna(0.0)
res["nuevo_saldo_1345_sum"] = res["nuevo_saldo_1345_sum"].fillna(0.0)
res["diferencia"]           = res["valor_cobro_sum"] - res["nuevo_saldo_1345_sum"]

tolerance = 0.01
conciliacion = res[res["diferencia"].abs() > tolerance].sort_values("_apto_key")

keys_cierre  = set(g_cierre["_apto_key"])
keys_balance = set(g_balance["_apto_key"])
only_cierre_keys  = sorted(keys_cierre - keys_balance)
only_balance_keys = sorted(keys_balance - keys_cierre)
both_keys         = sorted(keys_cierre & keys_balance)

solo_cierre_df  = res[res["_apto_key"].isin(only_cierre_keys)].sort_values("_apto_key")
solo_balance_df = res[res["_apto_key"].isin(only_balance_keys)].sort_values("_apto_key")

cobro_sin_saldo     = conciliacion[(conciliacion["valor_cobro_sum"] != 0) & (conciliacion["nuevo_saldo_1345_sum"] == 0)]
saldo_sin_cobro     = conciliacion[(conciliacion["valor_cobro_sum"] == 0) & (conciliacion["nuevo_saldo_1345_sum"] != 0)]
ambos_con_diferencia = conciliacion[(conciliacion["valor_cobro_sum"] != 0) & (conciliacion["nuevo_saldo_1345_sum"] != 0)]

n_diff = len(conciliacion)

# =========================
# PASO 5 — Panel de resultados
# =========================
st.markdown('<div class="section-card">', unsafe_allow_html=True)
section(5, "RESULTADOS DE LA CONCILIACIÓN", "📈")

# Resumen en lenguaje natural
if n_diff == 0:
    alert("Conciliación perfecta — no hay diferencias entre el Cierre y el Balance (1345*).", "success")
else:
    monto_total_diff = conciliacion["diferencia"].abs().sum()
    alert(
        f"{n_diff} apartamento{'s' if n_diff>1 else ''} con diferencias · "
        f"Monto total en discrepancia: {fmt_cop(monto_total_diff)}",
        "warning" if n_diff <= 10 else "error"
    )

# Métricas
st.markdown(f"""
<div class="metric-grid">
  <div class="metric-cell">
    <div class="metric-label">Aptos en Cierre</div>
    <div class="metric-value">{n_aptos_cierre}</div>
  </div>
  <div class="metric-cell">
    <div class="metric-label">Aptos en Balance 1345*</div>
    <div class="metric-value">{n_aptos_balance}</div>
  </div>
  <div class="metric-cell">
    <div class="metric-label">En ambos archivos</div>
    <div class="metric-value">{len(both_keys)}</div>
  </div>
  <div class="metric-cell {'danger' if n_diff > 0 else 'success'}">
    <div class="metric-label">Con diferencias</div>
    <div class="metric-value {'danger' if n_diff > 0 else 'success'}">{n_diff}</div>
  </div>
  <div class="metric-cell {'warning' if len(only_cierre_keys)>0 else ''}">
    <div class="metric-label">Solo en Cierre</div>
    <div class="metric-value {'warning' if len(only_cierre_keys)>0 else ''}">{len(only_cierre_keys)}</div>
  </div>
  <div class="metric-cell {'warning' if len(only_balance_keys)>0 else ''}">
    <div class="metric-label">Solo en Balance</div>
    <div class="metric-value {'warning' if len(only_balance_keys)>0 else ''}">{len(only_balance_keys)}</div>
  </div>
  <div class="metric-cell {'danger' if len(cobro_sin_saldo)>0 else ''}">
    <div class="metric-label">Cobro sin saldo 1345</div>
    <div class="metric-value {'danger' if len(cobro_sin_saldo)>0 else ''}">{len(cobro_sin_saldo)}</div>
  </div>
  <div class="metric-cell {'danger' if len(saldo_sin_cobro)>0 else ''}">
    <div class="metric-label">Saldo 1345 sin cobro</div>
    <div class="metric-value {'danger' if len(saldo_sin_cobro)>0 else ''}">{len(saldo_sin_cobro)}</div>
  </div>
</div>
""", unsafe_allow_html=True)

# Tablas
tabs = st.tabs([
    f"  Diferencias ({n_diff})  ",
    f"  Solo Cierre ({len(only_cierre_keys)})  ",
    f"  Solo Balance ({len(only_balance_keys)})  ",
    "  Todos los aptos  ",
    "  Agregado Cierre  ",
    "  Agregado Balance  ",
])

def style_diferencia(df_show: pd.DataFrame) -> pd.DataFrame:
    """Aplica colores a la columna Diferencia si existe."""
    if "Diferencia" not in df_show.columns:
        return df_show
    return df_show

with tabs[0]:
    if n_diff == 0:
        st.markdown('<div class="alert alert-success"><span>✓</span><span>Sin diferencias — la cartera está conciliada.</span></div>', unsafe_allow_html=True)
    else:
        disp = clean_display(conciliacion.reset_index(drop=True))
        # Formatear montos
        for col in ["Valor Cobro (Cierre)", "Saldo 1345 (Balance)", "Diferencia"]:
            if col in disp.columns:
                disp[col] = disp[col].apply(fmt_cop)
        st.dataframe(disp, use_container_width=True, height=400)

with tabs[1]:
    if len(only_cierre_keys) == 0:
        alert("Todos los apartamentos del Cierre tienen registro en el Balance.", "success")
    else:
        alert(f"{len(only_cierre_keys)} apartamentos cobrados sin contrapartida contable 1345*.", "warning")
        disp = clean_display(solo_cierre_df.reset_index(drop=True))
        for col in ["Valor Cobro (Cierre)", "Saldo 1345 (Balance)", "Diferencia"]:
            if col in disp.columns:
                disp[col] = disp[col].apply(fmt_cop)
        st.dataframe(disp, use_container_width=True, height=350)

with tabs[2]:
    if len(only_balance_keys) == 0:
        alert("Todos los apartamentos del Balance tienen registro en el Cierre.", "success")
    else:
        alert(f"{len(only_balance_keys)} apartamentos con saldo 1345* sin cobro registrado en el Cierre.", "warning")
        disp = clean_display(solo_balance_df.reset_index(drop=True))
        for col in ["Valor Cobro (Cierre)", "Saldo 1345 (Balance)", "Diferencia"]:
            if col in disp.columns:
                disp[col] = disp[col].apply(fmt_cop)
        st.dataframe(disp, use_container_width=True, height=350)

with tabs[3]:
    disp = clean_display(res.sort_values("_apto_key").reset_index(drop=True))
    for col in ["Valor Cobro (Cierre)", "Saldo 1345 (Balance)", "Diferencia"]:
        if col in disp.columns:
            disp[col] = disp[col].apply(fmt_cop)
    st.dataframe(disp, use_container_width=True, height=400)

with tabs[4]:
    disp = clean_display(g_cierre.sort_values("_apto_key").reset_index(drop=True))
    for col in ["Valor Cobro (Cierre)"]:
        if col in disp.columns:
            disp[col] = disp[col].apply(fmt_cop)
    st.dataframe(disp, use_container_width=True, height=400)

with tabs[5]:
    disp = clean_display(g_balance.sort_values("_apto_key").reset_index(drop=True))
    for col in ["Saldo 1345 (Balance)"]:
        if col in disp.columns:
            disp[col] = disp[col].apply(fmt_cop)
    st.dataframe(disp, use_container_width=True, height=400)

st.markdown('</div>', unsafe_allow_html=True)

# =========================
# Descargar
# =========================
st.markdown('<div class="section-card">', unsafe_allow_html=True)
section(6, "DESCARGAR RESULTADOS", "⬇")

def build_output_excel() -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as w:
        g_cierre.rename(columns={"_apto_key": "apto_key"}).to_excel(w, "agregado_cierre", index=False)
        g_balance.rename(columns={"_apto_key": "apto_key"}).to_excel(w, "agregado_balance", index=False)
        res.rename(columns={"_apto_key": "apto_key"}).to_excel(w, "todos_los_aptos", index=False)
        conciliacion.rename(columns={"_apto_key": "apto_key"}).to_excel(w, "diferencias", index=False)
        solo_cierre_df.rename(columns={"_apto_key": "apto_key"}).to_excel(w, "solo_cierre", index=False)
        solo_balance_df.rename(columns={"_apto_key": "apto_key"}).to_excel(w, "solo_balance", index=False)
    return output.getvalue()

col_dl, col_info = st.columns([1, 3])
with col_dl:
    st.download_button(
        "⬇  Descargar Excel",
        data=build_output_excel(),
        file_name="conciliacion_cartera.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
with col_info:
    st.caption("El archivo incluye 6 hojas: agregado_cierre · agregado_balance · todos_los_aptos · diferencias · solo_cierre · solo_balance")

st.markdown('</div>', unsafe_allow_html=True)
