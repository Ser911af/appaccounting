import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from io import BytesIO
from matplotlib.backends.backend_pdf import PdfPages

# Configuración de la página
st.set_page_config(page_title="DIAN Report Analyzer", layout="wide")

# Función para convertir a Excel
@st.cache_data
def convertir_a_excel(dataframe):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name="Resultados")
    return output.getvalue()

# Función para crear PDF
def crear_pdf(figures):
    pdf_output = BytesIO()
    with PdfPages(pdf_output) as pdf:
        for fig in figures:
            pdf.savefig(fig)
            plt.close(fig)
    return pdf_output.getvalue()

# Función para analizar datos
def analizar_datos(df):
    # Validar columnas necesarias
    required_columns = ["Fecha Emisión", "Total", "IVA", "Tipo de documento", "Grupo"]
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        st.error(f"El archivo no contiene las columnas requeridas: {', '.join(missing_columns)}")
        return None

    # Convertir 'Fecha Emisión' a formato de fecha
    df["Fecha Emisión"] = pd.to_datetime(df["Fecha Emisión"], format='%d-%m-%Y', errors="coerce")
    invalid_dates = df["Fecha Emisión"].isnull().sum()
    if invalid_dates > 0:
        st.warning(f"Se encontraron {invalid_dates} fechas mal formateadas que fueron ignoradas.")

    # Asegurarse de que las columnas 'Total' e 'IVA' son numéricas
    df["Total"] = pd.to_numeric(df["Total"], errors='coerce')
    df["IVA"] = pd.to_numeric(df["IVA"], errors='coerce')

    # Crear columna 'Base'
    df["Base"] = (df["IVA"].fillna(0)).round(0)

    # Extraer el nombre del mes
    month_mapping = {
        1: "January", 2: "February", 3: "March", 4: "April",
        5: "May", 6: "June", 7: "July", 8: "August",
        9: "September", 10: "October", 11: "November", 12: "December"
    }
    df["Mes"] = df["Fecha Emisión"].dt.month.map(month_mapping)

    return df

# Función para generar gráficos y tablas
def generar_resultados(df):
    meses_orden = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ]
    df["Mes"] = pd.Categorical(df["Mes"], categories=meses_orden, ordered=True)

    # Gráficos de IVA por mes
    iva_por_mes = df.groupby("Mes")["IVA"].sum().reindex(meses_orden, fill_value=0)
    iva_por_mes_millones = iva_por_mes / 1_000_000

    fig_iva, ax_iva = plt.subplots(figsize=(10, 6))
    ax_iva.plot(iva_por_mes_millones.index, iva_por_mes_millones.values, marker='o', linestyle='-', color='b')
    ax_iva.set_title("Evolución del IVA por Mes (en millones de pesos)", fontsize=16)
    ax_iva.set_xlabel("Mes", fontsize=12)
    ax_iva.set_ylabel("Total de IVA (Millones de Pesos)", fontsize=12)
    ax_iva.grid(True, linestyle='--', alpha=0.6)
    ax_iva.set_xticklabels(meses_orden, rotation=45)
    st.pyplot(fig_iva)

    # Tabla consolidada
    tipo_documentos = df["Tipo de documento"].unique()
    grados = ["Emitido", "Recibido"]
    tabla_resultados = []

    for tipo_doc in tipo_documentos:
        for grado in grados:
            df_filtro = df[(df["Tipo de documento"] == tipo_doc) & (df["Grupo"] == grado)]
            suma_por_mes = (
                df_filtro.groupby("Mes")["Base"].sum()
                .reindex(meses_orden, fill_value=0)
            )
            total_anual = suma_por_mes.sum()
            fila = [tipo_doc, grado] + list(suma_por_mes.values) + [total_anual]
            tabla_resultados.append(fila)

    columnas = ["Tipo Doc", "Grado"] + meses_orden + ["Total Anual"]
    tabla_df = pd.DataFrame(tabla_resultados, columns=columnas)
    st.markdown("### Tabla Consolidada")
    st.dataframe(tabla_df)

    # Descargar tabla consolidada
    excel_data = convertir_a_excel(tabla_df)
    st.download_button(
        label="Descargar tabla consolidada en Excel",
        data=excel_data,
        file_name="analisis_dian.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Menú de selección
st.sidebar.title("Menú")
option = st.sidebar.radio("Selecciona una funcionalidad:", ["IVA DIAN Report Analyzer", "DIAN Report Analyzer"])

# Subida del archivo
uploaded_file = st.file_uploader("Sube tu archivo Excel", type=["xlsx"])
if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        df = analizar_datos(df)
        if df is not None:
            if option == "IVA DIAN Report Analyzer":
                st.header("IVA DIAN Report Analyzer")
                generar_resultados(df)
            elif option == "DIAN Report Analyzer":
                st.header("DIAN Report Analyzer")
                generar_resultados(df)
    except Exception as e:
        st.error(f"Error procesando el archivo: {e}")
else:
    st.write("Por favor, sube un archivo Excel para comenzar.")
