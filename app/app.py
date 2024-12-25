import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from matplotlib.backends.backend_pdf import PdfPages

# Título de la aplicación
st.title("DIAN Report Analyzer")
st.subheader("Carga tu reporte DIAN y obtén un análisis detallado del archivo")

# Subida del archivo
uploaded_file = st.file_uploader("Sube tu archivo Excel", type=["xlsx"])

if uploaded_file:
    try:
        # Leer el archivo Excel
        df = pd.read_excel(uploaded_file)

        # Validar columnas necesarias
        required_columns = ["Fecha Emisión", "Total", "IVA", "Tipo de documento", "Grupo"]
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            st.error(f"El archivo no contiene las columnas requeridas: {', '.join(missing_columns)}")
        else:
            # Convertir 'Fecha Emisión' a formato de fecha y obtener el nombre del mes
            df["Fecha Emisión"] = pd.to_datetime(df["Fecha Emisión"], format='%d-%m-%Y', errors="coerce")
            df["Mes"] = df["Fecha Emisión"].dt.month_name(locale='es')

            # Validar fechas mal formateadas
            invalid_dates = df["Fecha Emisión"].isnull().sum()
            if invalid_dates > 0:
                st.warning(f"Se encontraron {invalid_dates} fechas mal formateadas que fueron ignoradas.")

            # Asegurarse de que las columnas 'Total' e 'IVA' son numéricas
            df["Total"] = pd.to_numeric(df["Total"], errors='coerce')
            df["IVA"] = pd.to_numeric(df["IVA"], errors='coerce')

            # Comprobar si hay valores nulos o inválidos
            if df["Total"].isnull().any() or df["IVA"].isnull().any():
                st.warning("Algunos valores de 'Total' o 'IVA' no son válidos y se han marcado como NaN.")

            # Crear columna 'Base' redondeando a enteros
            df["Base"] = (df["Total"].fillna(0) - df["IVA"].fillna(0)).round(0)

            # Filtrar valores nulos en 'Tipo de documento'
            df = df.dropna(subset=["Tipo de documento"])

            # Validar tipos de documentos inesperados
            expected_docs = {"Factura", "Nota Crédito", "Nota Débito"}
            tipo_documentos = df["Tipo de documento"].unique()
            unexpected_docs = set(tipo_documentos) - expected_docs
            if unexpected_docs:
                st.warning(f"Se encontraron tipos de documentos inesperados: {', '.join(unexpected_docs)}")

            # Calcular la tabla consolidada con pivot_table
            tabla_resultados = df.pivot_table(
                index=["Tipo de documento", "Grupo"],
                columns="Mes",
                values="Base",
                aggfunc="sum",
                fill_value=0
            ).reset_index()

            # Agregar columna de total anual
            tabla_resultados["Total Anual"] = tabla_resultados.iloc[:, 2:].sum(axis=1)

            # Mostrar tabla consolidada en la aplicación
            st.markdown("### Tabla consolidada:")
            st.dataframe(tabla_resultados)

            # Crear gráficos de barras por tipo de documento
            st.markdown("### Gráficos de barras: porcentaje relativo del valor por tipo de documento")

            for tipo_doc in tipo_documentos:
                tipo_df = tabla_resultados[tabla_resultados["Tipo de documento"] == tipo_doc]
                if tipo_df.empty:
                    st.warning(f"No hay datos suficientes para generar el gráfico de {tipo_doc}.")
                else:
                    fig, ax = plt.subplots(figsize=(10, 6))
                    total_anual = tipo_df["Total Anual"].sum()
                    valores_por_mes = tipo_df.iloc[:, 2:-1].sum().values
                    meses = tipo_df.columns[2:-1]

                    if total_anual > 0:
                        porcentajes = (valores_por_mes / total_anual) * 100
                        ax.bar(meses, porcentajes, color='skyblue', width=0.6)
                        ax.set_title(f"Porcentaje relativo de {tipo_doc}", fontsize=16)
                        ax.set_xlabel("Mes", fontsize=12)
                        ax.set_ylabel("Porcentaje (%)", fontsize=12)
                        ax.set_ylim(0, 100)
                        ax.set_xticklabels(meses, rotation=45)

                        # Etiquetas de porcentaje
                        for i, porcentaje in enumerate(porcentajes):
                            ax.text(i, porcentaje + 1, f"{porcentaje:.1f}%", ha='center', va='bottom', fontsize=10)

                        # Mostrar el gráfico
                        st.pyplot(fig)
                    else:
                        st.warning(f"No hay valores suficientes para el gráfico de {tipo_doc}.")

    except Exception as e:
        st.error(f"Error procesando el archivo: {e}")
else:
    st.write("Por favor, sube un archivo Excel para comenzar el análisis.")
