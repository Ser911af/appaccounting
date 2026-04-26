# tutorial.py — Mini tutorial visual con tarjetas paso a paso

import streamlit as st

STEPS = [
    {
        "num": "01",
        "icon": "📂",
        "head": "Carga los archivos",
        "body": (
            "Sube el Excel de <strong>Cierre</strong> y el de <strong>Balance</strong>. "
            "Puedes arrastrarlos o hacer clic en el área gris para buscarlos en tu equipo."
        ),
        "tip": "Ambos archivos deben estar en formato .xlsx",
    },
    {
        "num": "02",
        "icon": "🔍",
        "head": "Indica los encabezados",
        "body": (
            "Los archivos a veces tienen filas de título arriba. "
            "Mira la tabla de vista previa y escribe el <strong>número de fila</strong> "
            "donde aparecen los nombres de columna (Bloque, Código, Saldo...)."
        ),
        "tip": "Si los nombres están en la fila 1, deja el valor en 1",
    },
    {
        "num": "03",
        "icon": "📋",
        "head": "Columnas del Cierre",
        "body": (
            "La app detecta automáticamente las columnas de "
            "<strong>Bloque</strong>, <strong>Código de apartamento</strong> y "
            "<strong>Valor Cobro</strong>. "
            "Si la sugerencia no es correcta, cámbiala en el selector."
        ),
        "tip": "Revisa que el contador de aptos detectados sea correcto",
    },
    {
        "num": "04",
        "icon": "📊",
        "head": "Columnas del Balance",
        "body": (
            "Confirma la columna de <strong>clave de apartamento</strong> (ej: 1-101), "
            "la de <strong>Nuevo Saldo</strong> y la de <strong>Cuenta contable</strong>. "
            "Solo se usan las cuentas que empiezan con <strong>1345</strong>."
        ),
        "tip": "El sistema filtra automáticamente las cuentas 1345*",
    },
    {
        "num": "05",
        "icon": "📈",
        "head": "Lee los resultados",
        "body": (
            "Las métricas cambian en tiempo real. "
            "<strong style='color:#f85149'>Rojo</strong> = diferencia o problema. "
            "<strong style='color:#d29922'>Amarillo</strong> = apartamento en solo un archivo. "
            "<strong style='color:#3fb950'>Verde</strong> = todo en orden. "
            "Navega por las pestañas para ver el detalle."
        ),
        "tip": "La columna Diferencia se colorea automáticamente",
    },
    {
        "num": "06",
        "icon": "⬇️",
        "head": "Descarga el Excel",
        "body": (
            "El botón al final genera un archivo con <strong>6 hojas</strong>: "
            "diferencias, solo-cierre, solo-balance, todos los aptos y los agregados "
            "de cada archivo."
        ),
        "tip": "El archivo se llama conciliacion_cartera.xlsx",
    },
]


def render_tutorial() -> None:
    """Renderiza el tutorial desplegable con 6 tarjetas."""
    with st.expander("📖  ¿Cómo usar esta herramienta?  —  haz clic para ver la guía paso a paso"):
        cards_html = "\n".join(
            f"""<div class="tut-card">
                <div class="tut-num">{s["num"]}</div>
                <div class="tut-icon">{s["icon"]}</div>
                <div class="tut-head">{s["head"]}</div>
                <div class="tut-body">{s["body"]}</div>
                <div class="tut-tip">💡 {s["tip"]}</div>
            </div>"""
            for s in STEPS
        )
        st.markdown(
            f'<div class="tut-wrap">'
            f'<p class="tut-header">Guía rápida — 6 pasos para conciliar</p>'
            f'<div class="tut-grid">{cards_html}</div>'
            f'</div>',
            unsafe_allow_html=True,
        )
        st.caption(
            "Esta herramienta concilia la cartera de apartamentos comparando el Valor Cobro del Cierre "
            "contra el Nuevo Saldo de las cuentas 1345* del Balance. "
            "La clave de cruce es el número de apartamento en formato piso-número (ej: 2-9803)."
        )
