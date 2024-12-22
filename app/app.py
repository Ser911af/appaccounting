import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from io import BytesIO
from matplotlib.backends.backend_pdf import PdfPages
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Configurar nombres de los meses en español
month_mapping = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

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
            # Convertir 'Fecha Emisión' a formato de fecha
            df["Fecha Emisión"] = pd.to_datetime(df["Fecha Emisión"], format='%d-%m-%Y', errors="coerce")
            df["Total"] = pd.to_numeric(df["Total"], errors='coerce')
            df["IVA"] = pd.to_numeric(df["IVA"], errors='coerce')
            df["Base"] = (df["Total"].fillna(0) - df["IVA"].fillna(0)).round(0)
            df["Mes"] = df["Fecha Emisión"].dt.month.map(month_mapping)
            df["Mes"] = pd.Categorical(df["Mes"], categories=list(month_mapping.values()), ordered=True)

            tipo_documentos = df["Tipo de documento"].unique()
            grados = ["Emitido", "Recibido"]

            # Crear tabla consolidada
            tabla_resultados = []
            for tipo_doc in tipo_documentos:
                for grado in grados:
                    df_filtro = df[(df["Tipo de documento"] == tipo_doc) & (df["Grupo"] == grado)]
                    suma_por_mes = (
                        df_filtro.groupby("Mes")["Base"].sum()
                        .reindex(list(month_mapping.values()), fill_value=0)
                    )
                    total_anual = suma_por_mes.sum()
                    fila = [tipo_doc, grado] + list(suma_por_mes.values) + [total_anual]
                    tabla_resultados.append(fila)

            columnas = ["Tipo Doc", "Grado"] + list(month_mapping.values()) + ["Total Anual"]
            tabla_df = pd.DataFrame(tabla_resultados, columns=columnas)

            # Mostrar tabla en la aplicación
            st.markdown("### Tabla consolidada:")
            st.dataframe(tabla_df)

            # Crear gráficos de barras
            st.markdown("### Gráficos de barras:")
            all_figures = []
            for tipo_doc in tipo_documentos:
                fig, axes = plt.subplots(1, 2, figsize=(16, 6), sharey=True)
                fig.suptitle(f"Porcentaje relativo de {tipo_doc}", fontsize=16)
                for ax, grado in zip(axes, grados):
                    df_filtro = tabla_df[(tabla_df["Tipo Doc"] == tipo_doc) & (tabla_df["Grado"] == grado)]
                    total_anual = df_filtro["Total Anual"].values[0]
                    porcentajes = (df_filtro[list(month_mapping.values())].values.flatten() / total_anual) * 100
                    ax.bar(month_mapping.values(), porcentajes, color='skyblue', width=0.6)
                    ax.set_title(grado, fontsize=14)
                    ax.set_xlabel("Mes", fontsize=12)
                    ax.set_ylabel("Porcentaje (%)", fontsize=12)
                    ax.set_ylim(0, 100)
                    ax.set_xticklabels(month_mapping.values(), rotation=45)
                    for i, porcentaje in enumerate(porcentajes):
                        ax.text(i, porcentaje + 1, f"{porcentaje:.1f}%", ha='center', va='bottom', fontsize=10)
                st.pyplot(fig)
                all_figures.append(fig)

            # Descargar Excel con tabla y gráficos
            def generar_excel(tabla_df):
                output = BytesIO()
                wb = Workbook()
                ws = wb.active
                ws.title = "Resultados"
                for row in dataframe_to_rows(tabla_df, index=False, header=True):
                    ws.append(row)
                wb.save(output)
                return output.getvalue()

            st.markdown("### Descargar consolidado:")
            excel_data = generar_excel(tabla_df)
            st.download_button(
                label="Descargar Excel",
                data=excel_data,
                file_name="consolidado_dian.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error procesando el archivo: {e}")
else:
    st.write("Por favor, sube un archivo Excel para comenzar el análisis.")

