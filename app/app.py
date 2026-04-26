# -*- coding: utf-8 -*-
import io
import numpy as np
import pandas as pd
import streamlit as st

from styles   import DARK_CSS
from tutorial import render_tutorial
from utils    import (
    alert, build_table_from_row, clean_display, drop_all_empty_columns,
    find_col_fuzzy, fmt_cop, is_valid_apto_format, make_styled_table,
    normalize_apto_key, normalize_text, pick_amount_col,
    pick_apto_col_by_pattern, read_excel_cached, section, to_amount,
)

# ── Config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Conciliación de Cartera",
    layout="wide",
    initial_sidebar_state="collapsed",
)
st.markdown(DARK_CSS, unsafe_allow_html=True)

# ── Header ───────────────────────────────────────────────────────────────
st.markdown("""
<div class="app-header">
    <p class="app-title">Conciliación de Cartera</p>
    <p class="app-subtitle">Cierre vs Balance &nbsp;·&nbsp; cuentas 1345* &nbsp;·&nbsp; Frutto Foods</p>
</div>
""", unsafe_allow_html=True)

# ── Tutorial ─────────────────────────────────────────────────────────────
render_tutorial()

# ════════════════════════════════════════════════════════════════════════
# PASO 1 — Carga de archivos
# ════════════════════════════════════════════════════════════════════════
section(1, "CARGA LOS ARCHIVOS", "📂")

col_a, col_b = st.columns(2)
with col_a:
    cierre_file = st.file_uploader(
        "Archivo de Cierre (.xlsx)", type=["xlsx"], key="cierre",
        help="Reporte de Cierre exportado desde el sistema de cartera",
    )
with col_b:
    balance_file = st.file_uploader(
        "Archivo de Balance (.xlsx)", type=["xlsx"], key="balance",
        help="Balance de comprobación con cuentas 1345*",
    )

st.divider()

if not cierre_file or not balance_file:
    if cierre_file:
        alert("Cierre cargado ✓ — falta el archivo de Balance.", "warning")
    elif balance_file:
        alert("Balance cargado ✓ — falta el archivo de Cierre.", "warning")
    else:
        alert("Sube los dos archivos para comenzar la conciliación.", "info")
    st.stop()

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

alert(
    f"Archivos cargados — "
    f"Cierre: <strong>{cierre_file.name}</strong> ({len(raw_cierre)} filas) &nbsp;·&nbsp; "
    f"Balance: <strong>{balance_file.name}</strong> ({len(raw_balance)} filas)",
    "success",
)

# ════════════════════════════════════════════════════════════════════════
# PASO 2 — Fila de encabezados
# ════════════════════════════════════════════════════════════════════════
section(2, "INDICA LA FILA DE ENCABEZADOS", "🔍")
st.caption(
    "Cada archivo puede tener filas de título o logotipo arriba de los datos reales. "
    "Mira la vista previa y ajusta el número hasta que la primera fila muestre "
    "los nombres de columna."
)

tab_enc_c, tab_enc_b = st.tabs(["  📋  Cierre  ", "  📊  Balance  "])

with tab_enc_c:
    st.markdown("**Primeras filas tal como están en Excel — antes de aplicar encabezados:**")
    st.dataframe(
        raw_cierre.head(8).rename(columns=lambda x: f"col {x+1}"),
        use_container_width=True, height=200,
    )
    header_row_cierre_1based = st.number_input(
        "¿En qué fila están los NOMBRES de columna?",
        min_value=1, max_value=int(len(raw_cierre)),
        value=1, step=1, key="fila_header_cierre",
        help="Cuenta desde arriba. Si los nombres están en la primera fila, pon 1.",
    )
    header_row_cierre = int(header_row_cierre_1based) - 1
    df_cierre = build_table_from_row(raw_cierre, header_row_cierre)
    df_cierre  = drop_all_empty_columns(df_cierre)
    st.markdown("**Vista previa con encabezados aplicados — así quedan las columnas:**")
    st.dataframe(df_cierre.head(5), use_container_width=True, height=180)

with tab_enc_b:
    st.markdown("**Primeras filas tal como están en Excel — antes de aplicar encabezados:**")
    st.dataframe(
        raw_balance.head(8).rename(columns=lambda x: f"col {x+1}"),
        use_container_width=True, height=200,
    )
    header_row_balance_1based = st.number_input(
        "¿En qué fila están los NOMBRES de columna?",
        min_value=1, max_value=int(len(raw_balance)),
        value=1, step=1, key="fila_header_balance",
        help="Cuenta desde arriba. Si los nombres están en la primera fila, pon 1.",
    )
    header_row_balance = int(header_row_balance_1based) - 1
    df_balance = build_table_from_row(raw_balance, header_row_balance)
    df_balance  = drop_all_empty_columns(df_balance)
    st.markdown("**Vista previa con encabezados aplicados — así quedan las columnas:**")
    st.dataframe(df_balance.head(5), use_container_width=True, height=180)

st.divider()

# ════════════════════════════════════════════════════════════════════════
# PASO 3 — Columnas del Cierre
# ════════════════════════════════════════════════════════════════════════
section(3, "COLUMNAS DEL CIERRE", "📋")
st.caption("La app detecta las columnas automáticamente. Corrígelas si la sugerencia no es exacta.")

cols_cierre       = list(df_cierre.columns)
auto_codigo       = find_col_fuzzy(df_cierre, ["inmueble codigo", "inmueble código", "codigo", "código"]) or cols_cierre[0]
auto_bloque       = find_col_fuzzy(df_cierre, ["inmueble bloque", "bloque", "torre"]) or cols_cierre[0]
auto_valor_cobro  = pick_amount_col(df_cierre, ["valor cobro", "valor a cobrar", "cobro", "cuota", "facturado", "valor"]) or cols_cierre[0]

c1, c2, c3 = st.columns(3)
with c1:
    codigo_cierre_col = st.selectbox(
        "Número del apartamento", options=cols_cierre,
        index=cols_cierre.index(auto_codigo) if auto_codigo in cols_cierre else 0,
        help="Columna con el código del apartamento, ej: 9801, 203",
    )
with c2:
    bloque_cierre_col = st.selectbox(
        "Bloque / Torre", options=cols_cierre,
        index=cols_cierre.index(auto_bloque) if auto_bloque in cols_cierre else 0,
        help="Columna con el número de bloque o torre, ej: 1, 2, 3",
    )
with c3:
    valor_cobro_col = st.selectbox(
        "Valor Cobro", options=cols_cierre,
        index=cols_cierre.index(auto_valor_cobro) if auto_valor_cobro in cols_cierre else 0,
        help="Columna con el monto cobrado o facturado",
    )

# Construir clave y agregar Cierre
bloque_raw = df_cierre[bloque_cierre_col].astype(str).str.extract(r"(\d+)")[0]
codigo_raw = df_cierre[codigo_cierre_col].astype(str).str.extract(r"(\d{3,5})")[0]
piso = bloque_raw.str.lstrip("0").replace("", np.nan)

df_cierre["_apto_key_raw"] = np.where(piso.notna() & codigo_raw.notna(), piso + "-" + codigo_raw, np.nan)
df_cierre["_apto_key"]     = df_cierre["_apto_key_raw"].apply(normalize_apto_key)
df_cierre["_valor_cobro_num"] = to_amount(df_cierre[valor_cobro_col])
df_cierre_valid = df_cierre[df_cierre["_apto_key"].apply(is_valid_apto_format)].copy()

g_cierre = (
    df_cierre_valid.dropna(subset=["_apto_key"])
    .groupby("_apto_key", as_index=False)
    .agg(
        valor_cobro_sum=("_valor_cobro_num", "sum"),
        conteo_registros=(codigo_cierre_col, "count"),
    )
)

n_aptos_cierre = len(g_cierre)
if n_aptos_cierre == 0:
    alert(
        "No se detectaron apartamentos válidos. "
        "Revisa que las columnas <strong>Bloque</strong> y <strong>Número de apartamento</strong> sean correctas.",
        "error",
    )
else:
    total_cobro = g_cierre["valor_cobro_sum"].sum()
    alert(
        f"<strong>{n_aptos_cierre}</strong> apartamentos identificados en el Cierre &nbsp;·&nbsp; "
        f"Total cobro: <strong>{fmt_cop(total_cobro)}</strong>",
        "success",
    )

st.divider()

# ════════════════════════════════════════════════════════════════════════
# PASO 4 — Columnas del Balance
# ════════════════════════════════════════════════════════════════════════
section(4, "COLUMNAS DEL BALANCE", "📊")
st.caption("Solo se toman las filas cuya cuenta comience por 1345.")

cols_balance    = list(df_balance.columns)
auto_apto_bal   = (
    pick_apto_col_by_pattern(df_balance)
    or find_col_fuzzy(df_balance, ["apto", "apart", "unidad", "inmueble", "nit", "nombre nit"])
    or cols_balance[0]
)
auto_nuevo_saldo = (
    next((c for c in cols_balance if normalize_text(c) == "nuevo saldo"), None)
    or pick_amount_col(df_balance, ["nuevo saldo", "saldo nuevo", "saldo final", "saldo", "balance", "cartera"])
    or cols_balance[0]
)
auto_cuenta = find_col_fuzzy(df_balance, ["cuenta", "codigo", "código", "cod"]) or cols_balance[0]

b1, b2, b3 = st.columns(3)
with b1:
    apto_balance_col = st.selectbox(
        "Clave de apartamento", options=cols_balance,
        index=cols_balance.index(auto_apto_bal) if auto_apto_bal in cols_balance else 0,
        help="Columna con clave tipo 1-101, 2-203, etc.",
    )
with b2:
    nuevo_saldo_col = st.selectbox(
        "Nuevo Saldo", options=cols_balance,
        index=cols_balance.index(auto_nuevo_saldo) if auto_nuevo_saldo in cols_balance else 0,
        help="Columna con el saldo contable del apartamento",
    )
with b3:
    cuenta_col = st.selectbox(
        "Cuenta contable", options=cols_balance,
        index=cols_balance.index(auto_cuenta) if auto_cuenta in cols_balance else 0,
        help="Se filtran las que empiezan con 1345",
    )

# Procesar Balance
df_balance["_apto_key"]            = df_balance[apto_balance_col].apply(normalize_apto_key)
df_balance["_apto_key_valid"]      = df_balance["_apto_key"].apply(is_valid_apto_format)
df_balance["_nuevo_saldo_num"]     = to_amount(df_balance[nuevo_saldo_col])
df_balance["_is_1345"]             = df_balance[cuenta_col].astype(str).str.strip().str.startswith("1345")
df_balance["_nuevo_saldo_1345_num"] = np.where(df_balance["_is_1345"], df_balance["_nuevo_saldo_num"], 0.0)
df_balance_valid = df_balance[df_balance["_apto_key_valid"]].copy()

g_balance = (
    df_balance_valid.dropna(subset=["_apto_key"])
    .groupby("_apto_key", as_index=False)
    .agg(
        nuevo_saldo_1345_sum=("_nuevo_saldo_1345_num", "sum"),
        conteo_registros=(cuenta_col, "count"),
    )
)

n_aptos_balance = len(g_balance)
n_filas_1345    = int(df_balance["_is_1345"].sum())
if n_aptos_balance == 0:
    alert(
        "No se detectaron apartamentos válidos en el Balance. "
        "Revisa la columna <strong>Clave de apartamento</strong>.",
        "error",
    )
else:
    total_saldo = g_balance["nuevo_saldo_1345_sum"].sum()
    alert(
        f"<strong>{n_aptos_balance}</strong> apartamentos identificados en el Balance &nbsp;·&nbsp; "
        f"{n_filas_1345} filas con cuentas 1345* &nbsp;·&nbsp; "
        f"Total saldo: <strong>{fmt_cop(total_saldo)}</strong>",
        "success",
    )

st.divider()

# Guard: sin aptos no hay conciliación útil
if n_aptos_cierre == 0 or n_aptos_balance == 0:
    alert("Corrige las columnas en los pasos anteriores para ver la conciliación.", "warning")
    st.stop()

# ════════════════════════════════════════════════════════════════════════
# PASO 5 — Conciliación
# ════════════════════════════════════════════════════════════════════════
res = pd.merge(g_cierre, g_balance, on="_apto_key", how="outer")
res["valor_cobro_sum"]      = res["valor_cobro_sum"].fillna(0.0)
res["nuevo_saldo_1345_sum"] = res["nuevo_saldo_1345_sum"].fillna(0.0)
res["diferencia"]           = res["valor_cobro_sum"] - res["nuevo_saldo_1345_sum"]

TOLERANCE = 0.01
conciliacion = res[res["diferencia"].abs() > TOLERANCE].sort_values("_apto_key")

keys_cierre  = set(g_cierre["_apto_key"])
keys_balance = set(g_balance["_apto_key"])
only_cierre_keys  = sorted(keys_cierre - keys_balance)
only_balance_keys = sorted(keys_balance - keys_cierre)
both_keys         = sorted(keys_cierre & keys_balance)

solo_cierre_df  = res[res["_apto_key"].isin(only_cierre_keys)].sort_values("_apto_key")
solo_balance_df = res[res["_apto_key"].isin(only_balance_keys)].sort_values("_apto_key")

cobro_sin_saldo      = conciliacion[(conciliacion["valor_cobro_sum"] != 0) & (conciliacion["nuevo_saldo_1345_sum"] == 0)]
saldo_sin_cobro      = conciliacion[(conciliacion["valor_cobro_sum"] == 0) & (conciliacion["nuevo_saldo_1345_sum"] != 0)]
ambos_con_diferencia = conciliacion[(conciliacion["valor_cobro_sum"] != 0) & (conciliacion["nuevo_saldo_1345_sum"] != 0)]
n_diff = len(conciliacion)

# ── Panel de resultados ──────────────────────────────────────────────────
section(5, "RESULTADOS DE LA CONCILIACIÓN", "📈")

if n_diff == 0:
    alert("Conciliación perfecta — no hay diferencias entre el Cierre y el Balance (1345*).", "success")
else:
    monto_total = conciliacion["diferencia"].abs().sum()
    kind = "warning" if n_diff <= 10 else "error"
    alert(
        f"<strong>{n_diff}</strong> apartamento{'s' if n_diff > 1 else ''} con diferencias &nbsp;·&nbsp; "
        f"Monto total en discrepancia: <strong>{fmt_cop(monto_total)}</strong>",
        kind,
    )

# Métricas visuales
def mcell(label, value, css_class=""):
    cls = f"metric-cell {css_class}".strip()
    val_cls = f"metric-value {css_class}".strip()
    return (
        f'<div class="{cls}">'
        f'<div class="metric-label">{label}</div>'
        f'<div class="{val_cls}">{value}</div>'
        f'</div>'
    )

st.markdown(
    f'<div class="metric-grid">'
    + mcell("Aptos en Cierre",         n_aptos_cierre)
    + mcell("Aptos en Balance 1345*",  n_aptos_balance)
    + mcell("En ambos archivos",        len(both_keys))
    + mcell("Con diferencias",          n_diff,                "danger" if n_diff > 0 else "success")
    + mcell("Solo en Cierre",           len(only_cierre_keys),  "warning" if only_cierre_keys else "")
    + mcell("Solo en Balance",          len(only_balance_keys), "warning" if only_balance_keys else "")
    + mcell("Cobro sin saldo 1345",     len(cobro_sin_saldo),   "danger" if cobro_sin_saldo.shape[0] > 0 else "")
    + mcell("Saldo 1345 sin cobro",     len(saldo_sin_cobro),   "danger" if saldo_sin_cobro.shape[0] > 0 else "")
    + '</div>',
    unsafe_allow_html=True,
)

# ── Tablas en pestañas ───────────────────────────────────────────────────
tabs = st.tabs([
    f"  ⚠  Diferencias ({n_diff})  ",
    f"  ← Solo Cierre ({len(only_cierre_keys)})  ",
    f"  → Solo Balance ({len(only_balance_keys)})  ",
    "  ≡  Todos los aptos  ",
    "  Agregado Cierre  ",
    "  Agregado Balance  ",
])

TABLE_HEIGHT = 420

with tabs[0]:
    if n_diff == 0:
        alert("Sin diferencias — la cartera está conciliada.", "success")
    else:
        # Filtro de monto mínimo
        max_diff = float(conciliacion["diferencia"].abs().max())
        min_filter = st.slider(
            "Mostrar diferencias mayores a:", min_value=0.0,
            max_value=float(max_diff), value=0.0, step=1000.0,
            format="$ %,.0f",
            help="Filtra las diferencias pequeñas para enfocarte en las más importantes",
        )
        filtered = conciliacion[conciliacion["diferencia"].abs() >= min_filter]
        if len(filtered) == 0:
            alert("Ninguna diferencia supera ese umbral.", "info")
        else:
            st.caption(f"Mostrando {len(filtered)} de {n_diff} apartamentos")
            st.dataframe(make_styled_table(filtered), use_container_width=True, height=TABLE_HEIGHT)

with tabs[1]:
    if not only_cierre_keys:
        alert("Todos los apartamentos del Cierre tienen registro en el Balance.", "success")
    else:
        alert(f"{len(only_cierre_keys)} apartamentos cobrados sin contrapartida contable 1345*.", "warning")
        st.dataframe(make_styled_table(solo_cierre_df), use_container_width=True, height=TABLE_HEIGHT)

with tabs[2]:
    if not only_balance_keys:
        alert("Todos los apartamentos del Balance tienen registro en el Cierre.", "success")
    else:
        alert(f"{len(only_balance_keys)} apartamentos con saldo 1345* sin cobro en el Cierre.", "warning")
        st.dataframe(make_styled_table(solo_balance_df), use_container_width=True, height=TABLE_HEIGHT)

with tabs[3]:
    st.dataframe(
        make_styled_table(res.sort_values("_apto_key")),
        use_container_width=True, height=TABLE_HEIGHT,
    )

with tabs[4]:
    disp = clean_display(g_cierre.sort_values("_apto_key").reset_index(drop=True))
    disp_s = disp.style.format(fmt_cop, subset=["Valor Cobro (Cierre)"], na_rep="—")
    st.dataframe(disp_s, use_container_width=True, height=TABLE_HEIGHT)

with tabs[5]:
    disp = clean_display(g_balance.sort_values("_apto_key").reset_index(drop=True))
    disp_s = disp.style.format(fmt_cop, subset=["Saldo 1345 (Balance)"], na_rep="—")
    st.dataframe(disp_s, use_container_width=True, height=TABLE_HEIGHT)

st.divider()

# ════════════════════════════════════════════════════════════════════════
# PASO 6 — Descarga
# ════════════════════════════════════════════════════════════════════════
section(6, "DESCARGA LOS RESULTADOS", "⬇")


def build_output_excel() -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as w:
        g_cierre.rename(columns={"_apto_key": "apto_key"}).to_excel(w, "agregado_cierre",   index=False)
        g_balance.rename(columns={"_apto_key": "apto_key"}).to_excel(w, "agregado_balance",  index=False)
        res.rename(columns={"_apto_key": "apto_key"}).to_excel(w,         "todos_los_aptos",  index=False)
        conciliacion.rename(columns={"_apto_key": "apto_key"}).to_excel(w,"diferencias",       index=False)
        solo_cierre_df.rename(columns={"_apto_key": "apto_key"}).to_excel(w, "solo_cierre",   index=False)
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
    st.caption(
        "El archivo incluye 6 hojas: "
        "agregado_cierre · agregado_balance · todos_los_aptos · diferencias · solo_cierre · solo_balance"
    )
