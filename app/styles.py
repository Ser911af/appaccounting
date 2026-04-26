# styles.py — CSS del tema oscuro para la app de conciliación

DARK_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

.stApp { background-color: #0d1117; color: #e6edf3; }

/* ── Header ── */
.app-header {
    background: linear-gradient(135deg, #161b22 0%, #0d1117 100%);
    border-bottom: 1px solid #21262d;
    padding: 2rem 2.5rem 1.5rem;
    margin: -1rem -1rem 2rem -1rem;
}
.app-title {
    font-size: 1.6rem; font-weight: 700; color: #f0f6fc;
    letter-spacing: -0.02em; margin: 0 0 0.25rem 0;
}
.app-subtitle { font-size: 0.85rem; color: #8b949e; font-weight: 400; margin: 0; }

/* ── Título de sección ── */
.section-title {
    font-size: 0.7rem; font-weight: 600; letter-spacing: 0.1em;
    text-transform: uppercase; color: #58a6ff;
    margin: 0 0 0.8rem 0;
    display: flex; align-items: center; gap: 0.5rem;
    padding-bottom: 0.6rem;
    border-bottom: 1px solid #21262d;
}
.step-badge {
    background: #1f6feb; color: #fff; border-radius: 50%;
    width: 20px; height: 20px;
    display: inline-flex; align-items: center; justify-content: center;
    font-size: 0.65rem; font-weight: 700; flex-shrink: 0;
}

/* ── Tutorial cards ── */
.tut-wrap {
    background: #0d1117;
    border: 1px solid #21262d;
    border-radius: 10px;
    padding: 1.25rem;
    margin-bottom: 1.5rem;
}
.tut-header {
    font-size: 0.72rem; font-weight: 600; letter-spacing: 0.08em;
    text-transform: uppercase; color: #8b949e;
    margin: 0 0 1rem 0;
}
.tut-grid {
    display: grid;
    grid-template-columns: repeat(6, 1fr);
    gap: 8px;
}
@media (max-width: 1100px) { .tut-grid { grid-template-columns: repeat(3, 1fr); } }
@media (max-width: 700px)  { .tut-grid { grid-template-columns: repeat(2, 1fr); } }

.tut-card {
    background: #161b22;
    border: 1px solid #21262d;
    border-radius: 8px;
    padding: 1rem 0.85rem;
    display: flex; flex-direction: column; gap: 0.35rem;
    transition: border-color 0.15s;
}
.tut-card:hover { border-color: #388bfd; }
.tut-num {
    font-size: 0.65rem; font-weight: 700; color: #1f6feb;
    letter-spacing: 0.1em;
}
.tut-icon { font-size: 1.4rem; line-height: 1; }
.tut-head { font-size: 0.78rem; font-weight: 600; color: #f0f6fc; margin-top: 0.15rem; }
.tut-body { font-size: 0.73rem; color: #8b949e; line-height: 1.5; }
.tut-tip {
    font-size: 0.68rem; color: #3fb950;
    margin-top: 0.4rem; padding-top: 0.4rem;
    border-top: 1px solid #21262d;
}

/* ── Métricas ── */
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
.metric-cell          { background: #161b22; padding: 1rem 1.25rem; text-align: center; }
.metric-cell.danger   { background: #1a0f0f; }
.metric-cell.warning  { background: #141008; }
.metric-cell.success  { background: #0a1a0a; }
.metric-label {
    font-size: 0.68rem; font-weight: 500; color: #8b949e;
    text-transform: uppercase; letter-spacing: 0.06em; margin-bottom: 0.4rem;
}
.metric-value         { font-size: 1.8rem; font-weight: 700; color: #f0f6fc; line-height: 1; }
.metric-value.danger  { color: #f85149; }
.metric-value.warning { color: #d29922; }
.metric-value.success { color: #3fb950; }

/* ── Alertas inline ── */
.alert {
    padding: 0.7rem 1rem; border-radius: 6px;
    font-size: 0.82rem; margin: 0.6rem 0;
    display: flex; align-items: flex-start; gap: 0.6rem;
}
.alert-success { background: #0a1a0a; border-left: 3px solid #3fb950; color: #3fb950; }
.alert-warning { background: #1a1500; border-left: 3px solid #d29922; color: #d29922; }
.alert-error   { background: #1a0a0a; border-left: 3px solid #f85149; color: #f85149; }
.alert-info    { background: #0d1f3c; border-left: 3px solid #58a6ff; color: #58a6ff; }

/* ── Streamlit overrides ── */
.stFileUploader > div {
    background: #161b22 !important;
    border: 1px dashed #30363d !important;
    border-radius: 8px !important;
}
.stFileUploader > div:hover { border-color: #58a6ff !important; }
div[data-testid="stFileUploadDropzone"] { background: #161b22 !important; }

.stNumberInput > div > div > input,
.stSelectbox > div > div > div {
    background-color: #161b22 !important;
    border: 1px solid #30363d !important;
    color: #e6edf3 !important;
    border-radius: 6px !important;
}
.stSelectbox > div > div > div:hover { border-color: #58a6ff !important; }

.stTabs [data-baseweb="tab-list"] {
    background: #161b22; border-bottom: 1px solid #21262d; gap: 0;
}
.stTabs [data-baseweb="tab"] {
    background: transparent; color: #8b949e;
    font-size: 0.8rem; font-weight: 500;
    padding: 0.6rem 1rem;
    border-bottom: 2px solid transparent;
}
.stTabs [aria-selected="true"] {
    color: #58a6ff !important;
    border-bottom: 2px solid #58a6ff !important;
    background: transparent !important;
}

.stDownloadButton > button {
    background: #1f6feb !important; color: #fff !important;
    border: none !important; border-radius: 6px !important;
    font-weight: 600 !important; font-size: 0.85rem !important;
    padding: 0.6rem 1.5rem !important;
}
.stDownloadButton > button:hover { background: #388bfd !important; }

.stButton > button {
    background: #21262d !important; color: #e6edf3 !important;
    border: 1px solid #30363d !important; border-radius: 6px !important;
    font-size: 0.82rem !important;
}

/* Expander */
.streamlit-expanderHeader {
    background: #161b22 !important;
    border: 1px solid #21262d !important;
    border-radius: 8px !important;
    color: #8b949e !important;
    font-size: 0.8rem !important;
    font-weight: 600 !important;
}
.streamlit-expanderContent {
    background: #0d1117 !important;
    border: 1px solid #21262d !important;
    border-top: none !important;
}

/* Labels */
.stSelectbox label, .stNumberInput label, .stFileUploader label {
    color: #8b949e !important; font-size: 0.78rem !important;
    font-weight: 500 !important; text-transform: uppercase !important;
    letter-spacing: 0.06em !important;
}
h2, h3 { color: #f0f6fc !important; }
h4 { color: #8b949e !important; font-weight: 500 !important; font-size: 0.85rem !important; }
.stCaption { color: #6e7681 !important; font-size: 0.75rem !important; }
.stDataFrame { border-radius: 6px; overflow: hidden; }
[data-testid="stDataFrameResizable"] {
    border: 1px solid #21262d !important; border-radius: 6px !important;
}
section[data-testid="stSidebar"] {
    background: #161b22; border-right: 1px solid #21262d;
}
</style>
"""
