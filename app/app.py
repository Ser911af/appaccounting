import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from io import BytesIO
from matplotlib.backends.backend_pdf import PdfPages
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import BarChart, Reference

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

            # Configurar nombres de meses
            month_mapping = {
                1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
                5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
                9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
            }
            df["Mes"] = df["Fecha Emisión"].dt.month.map(month_mapping)
            meses_orden = list(month_mapping.values())
            df["Mes"] = pd.Categorical(df["Mes"], categories=meses_orden, ordered=True)

            # Obtener valores únicos
            tipo_documentos = df["Tipo de documento"].unique()
            grados = ["Emitido", "Recibido"]

            # Crear tabla consolidada
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
            tabla_df = pd.DataFrame(tabla_resultados, columns=columnas).round(0)

            # Mostrar tabla en la aplicación
            st.markdown("### Tabla consolidada:")
            st.dataframe(tabla_df)

            # Crear gráficos en la interfaz
            st.markdown("### Gráficos de barras: porcentaje relativo del valor por tipo de documento")
            all_figures = []
            for tipo_doc in tipo_documentos:
                fig, axes = plt.subplots(1, 2, figsize=(16, 6), sharey=True)
                fig.suptitle(f"Porcentaje relativo de {tipo_doc}", fontsize=16)

                for ax, grado in zip(axes, grados):
                    df_filtro = tabla_df[(tabla_df["Tipo Doc"] == tipo_doc) & (tabla_df["Grado"] == grado)]
                    total_anual = df_filtro["Total Anual"].values[0]
                    porcentajes = (df_filtro[meses_orden].values.flatten() / total_anual) * 100

                    ax.bar(meses_orden, porcentajes, color='skyblue', width=0.6)
                    ax.set_title(grado, fontsize=14)
                    ax.set_xlabel("Mes", fontsize=12)
                    ax.set_ylabel("Porcentaje (%)", fontsize=12)
                    ax.set_ylim(0, 100)
                    ax.set_xticks(range(len(meses_orden)))
                    ax.set_xticklabels(meses_orden, rotation=45)

                    for i, porcentaje in enumerate(porcentajes):
                        ax.text(i, porcentaje + 1, f"{porcentaje:.1f}%", ha='center', va='bottom', fontsize=10)

                st.pyplot(fig)
                all_figures.append(fig)

            # Crear PDF
            def crear_pdf(figures):
                pdf_output = BytesIO()
                with PdfPages(pdf_output) as pdf:
                    for fig in figures:
                        pdf.savefig(fig)
                        plt.close(fig)
                return pdf_output.getvalue()

            pdf_data = crear_pdf(all_figures)
            st.markdown("### Descargar gráficos en PDF")
            st.download_button(
                label="Descargar gráficos en PDF",
                data=pdf_data,
                file_name="gráficos_dian.pdf",
                mime="application/pdf"
            )

            # Crear Excel con gráficos
            def generar_excel_con_graficos(dataframe):
                output = BytesIO()
                wb = Workbook()
                ws = wb.active
                ws.title = "Resultados"

                for row in dataframe_to_rows(dataframe, index=False, header=True):
                    ws.append(row)

                for i, tipo_doc in enumerate(tipo_documentos, start=1):
                    chart = BarChart()
                    chart.title = f"Porcentaje relativo de {tipo_doc}"
                    chart.y_axis.title = "Porcentaje (%)"
                    chart.x_axis.title = "Mes"

                    data = Reference(ws, min_col=3, max_col=14, min_row=2 + (i - 1) * 2, max_row=2 + (i - 1) * 2)
                    categories = Reference(ws, min_col=3, max_col=14, min_row=1)
                    chart.add_data(data, titles_from_data=False)
                    chart.set_categories(categories)
                    ws.add_chart(chart, f"N{10 * i}")

                wb.save(output)
                return output.getvalue()

            excel_data = generar_excel_con_graficos(tabla_df)
            st.download_button(
                label="Descargar tabla y gráficos en Excel",
                data=excel_data,
                file_name="consolidado_dian.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error procesando el archivo: {e}")
else:
    st.write("Por favor, sube un archivo Excel para comenzar el análisis.")
