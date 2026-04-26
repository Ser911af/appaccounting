# utils.py — funciones de limpieza, detección de columnas y formateo

import io
import re
import unicodedata
from typing import Optional, List

import numpy as np
import pandas as pd
import streamlit as st

# ── Patrón de clave de apartamento ──────────────────────────────────────
APT_SEP_REGEX   = r"[-_/\.]"
PI_SO_NUM_FULL  = re.compile(rf"^\s*(\d{{1,2}})\s*{APT_SEP_REGEX}\s*(\d{{3,5}})\s*$")

# ────────────────────────────────────────────────────────────────────────
# Texto y limpieza
# ────────────────────────────────────────────────────────────────────────

def normalize_text(s) -> str:
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
    return df.loc[:, tmp.notna().any(axis=0).values]


def to_amount(series: pd.Series) -> pd.Series:
    """Convierte texto numérico (formato colombiano 1.234,56) a float."""
    s = series.fillna("").astype(str)
    s = s.replace({r'[^0-9\-,\.]': ''}, regex=True)
    s = s.str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
    return pd.to_numeric(s, errors='coerce')


def fmt_cop(val) -> str:
    """Formatea número como moneda colombiana: $ 1.234.567"""
    try:
        v = float(val)
        if pd.isna(v):
            return "—"
        sign = "-" if v < 0 else ""
        formatted = f"{abs(v):,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return f"{sign}$ {formatted}"
    except Exception:
        return str(val)


# ────────────────────────────────────────────────────────────────────────
# Construcción de tabla desde fila de encabezados
# ────────────────────────────────────────────────────────────────────────

def build_table_from_row(df_raw: pd.DataFrame, header_row_idx: int) -> pd.DataFrame:
    headers = df_raw.iloc[header_row_idx].astype(str).tolist()
    headers_norm = []
    for i, h in enumerate(headers):
        h_norm = normalize_text(h)
        headers_norm.append(f"col_{i}" if h_norm in ("", "unnamed: 0", "nan") else h)
    seen, unique_headers = {}, []
    for h in headers_norm:
        if h not in seen:
            seen[h] = 0
            unique_headers.append(h)
        else:
            seen[h] += 1
            unique_headers.append(f"{h}_{seen[h]}")
    data = df_raw.iloc[header_row_idx + 1:].copy()
    data.columns = unique_headers
    data = data.dropna(how="all")
    for c in list(data.select_dtypes(include=["object"]).columns):
        data[c] = data[c].astype(str).str.strip()
    return data


# ────────────────────────────────────────────────────────────────────────
# Clave de apartamento
# ────────────────────────────────────────────────────────────────────────

def normalize_apto_key(s) -> Optional[str]:
    if pd.isna(s):
        return None
    m = PI_SO_NUM_FULL.match(str(s).strip())
    if not m:
        return None
    return f"{int(m.group(1))}-{int(m.group(2))}"


def is_valid_apto_format(s) -> bool:
    if pd.isna(s):
        return False
    return bool(PI_SO_NUM_FULL.match(str(s).strip()))


# ────────────────────────────────────────────────────────────────────────
# Detección automática de columnas
# ────────────────────────────────────────────────────────────────────────

def find_col_fuzzy(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    cols  = list(df.columns)
    ncols = [normalize_text(c) for c in cols]
    best, best_score = None, 0
    for i, nc in enumerate(ncols):
        score = sum(kw in nc for kw in candidates)
        if score > best_score:
            best, best_score = cols[i], score
    return best  # None si ninguna columna hace match


def pick_apto_col_by_pattern(df: pd.DataFrame) -> Optional[str]:
    best_col, best_hits = None, -1
    sample_n = min(300, len(df))
    for col in df.columns:
        hits = df[col].dropna().astype(str).head(sample_n).apply(
            lambda x: 1 if is_valid_apto_format(x) else 0
        ).sum()
        if hits > best_hits:
            best_hits, best_col = hits, col
    return best_col if best_hits > 0 else None


def pick_amount_col(df: pd.DataFrame, prefer_keywords: List[str]) -> Optional[str]:
    kw_col = find_col_fuzzy(df, prefer_keywords)
    if kw_col is not None:
        return kw_col
    best_col, best_sum = None, -1
    for col in df.columns:
        s = np.nansum(np.abs(to_amount(df[col]).values))
        if np.isfinite(s) and s > best_sum:
            best_sum, best_col = s, col
    return best_col


# ────────────────────────────────────────────────────────────────────────
# Limpieza de tablas para mostrar al usuario
# ────────────────────────────────────────────────────────────────────────

RENAME_MAP = {
    "_apto_key":            "Apartamento",
    "valor_cobro_sum":      "Valor Cobro (Cierre)",
    "nuevo_saldo_1345_sum": "Saldo 1345 (Balance)",
    "diferencia":           "Diferencia",
    "conteo_registros_x":   "Registros Cierre",
    "conteo_registros_y":   "Registros Balance",
    "conteo_registros":     "Registros",
}


def clean_display(df: pd.DataFrame) -> pd.DataFrame:
    """Renombra columnas conocidas y elimina las internas (_*)."""
    result = df.rename(columns={k: v for k, v in RENAME_MAP.items() if k in df.columns})
    drop   = [c for c in result.columns if c.startswith("_")]
    return result.drop(columns=drop, errors="ignore")


def make_styled_table(df: pd.DataFrame) -> "pd.io.formats.style.Styler":
    """
    Limpia columnas, colorea Diferencia (rojo negativo / amarillo positivo / verde cero)
    y formatea montos como moneda colombiana.
    """
    disp = clean_display(df.reset_index(drop=True))

    def diff_color(val):
        try:
            v = float(val)
            if abs(v) < 0.01:
                return "color: #3fb950; font-weight: 600"
            elif v > 0:
                return "color: #d29922; font-weight: 600"
            else:
                return "color: #f85149; font-weight: 600"
        except Exception:
            return ""

    styled = disp.style

    if "Diferencia" in disp.columns:
        styled = styled.map(diff_color, subset=["Diferencia"])

    money_cols = [c for c in disp.columns if any(k in c for k in ["Cobro", "Saldo", "Diferencia"])]
    for col in money_cols:
        styled = styled.format(fmt_cop, subset=[col], na_rep="—")

    return styled


# ────────────────────────────────────────────────────────────────────────
# Lectura con caché
# ────────────────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def read_excel_cached(file_bytes: bytes, filename: str) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(file_bytes), header=None, dtype=str)


# ────────────────────────────────────────────────────────────────────────
# Helpers de UI
# ────────────────────────────────────────────────────────────────────────

def section(number: int, title: str, icon: str = "") -> None:
    st.markdown(
        f'<div class="section-title">'
        f'<span class="step-badge">{number}</span>'
        f'{icon}&nbsp;{title}'
        f'</div>',
        unsafe_allow_html=True,
    )


def alert(msg: str, kind: str = "info") -> None:
    icons = {"success": "✓", "warning": "⚠", "error": "✕", "info": "ℹ"}
    st.markdown(
        f'<div class="alert alert-{kind}">'
        f'<span>{icons.get(kind, "ℹ")}</span><span>{msg}</span>'
        f'</div>',
        unsafe_allow_html=True,
    )
