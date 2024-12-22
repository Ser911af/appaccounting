import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from io import BytesIO
from matplotlib.backends.backend_pdf import PdfPages
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.utils.dataframe import dataframe_to_rows

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
            invalid_dates = df["Fecha Emisión"].isnull().sum()
            if invalid_dates > 0:
                st.warning(f"Se encontraron {invalid_dates} fechas mal formateadas que fueron ignoradas.")
            else:
                st.success("Todas las fechas fueron convertidas correctamente.")

            # Asegurarse de que las columnas 'Total' e 'IVA' son numéricas
            df["Total"] = pd.to_numeric(df["Total"], errors='coerce')
            df["IVA"] = pd.to_numeric(df["IVA"], errors='coerce')

            # Comprobar si hay valores nulos o inválidos
            if df["Total"].isnull().any() or df["IVA"].isnull().any():
                st.warning("Algunos valores de 'Total' o 'IVA' no son válidos y se han marcado como NaN.")

            # Crear columna 'Base' redondeando a enteros
            df["Base"] = (df["Total"].fillna(0) - df["IVA"].fillna(0)).round(0)

            # Extraer el nombre del mes de forma manual
            month_mapping = {
                1: "January", 2: "February", 3: "March", 4: "April",
                5: "May", 6: "June", 7: "July", 8: "August",
                9: "September", 10: "October", 11: "November", 12: "December"
            }
            df["Mes"] = df["Fecha Emisión"].dt.month.map(month_mapping)

            # Ordenar meses correctamente
            meses_orden = list(month_mapping.values())
            df["Mes"] = pd.Categorical(df["Mes"], categories=meses_orden, ordered=True)

            # Obtener valores únicos de 'Tipo de documento' y 'Grupo'
            tipo_documentos = df["Tipo de documento"].unique()
            grados = ["Emitido", "Recibido"]

            # Crear tabla consolidada
            tabla_resultados = []

            for tipo_doc in tipo_documentos:
                for grado in grados:
                    # Filtrar datos por 'Tipo de documento' y 'Grupo'
                    df_filtro = df[(df["Tipo de documento"] == tipo_doc) & (df["Grupo"] == grado)]

                    # Sumar 'Base' por mes
                    suma_por_mes = (
                        df_filtro.groupby("Mes")["Base"].sum()
                        .reindex(meses_orden, fill_value=0)
                    )

                    # Calcular total anual
                    total_anual = suma_por_mes.sum()

                    # Crear fila de resultados
                    fila = [tipo_doc, grado] + list(suma_por_mes.values) + [total_anual]
                    tabla_resultados.append(fila)

            # Crear DataFrame con la tabla consolidada
            columnas = ["Tipo Doc", "Grado"] + meses_orden + ["Total Anual"]
            tabla_df = pd.DataFrame(tabla_resultados, columns=columnas)

            # Redondear valores a enteros
            tabla_df = tabla_df.round(0)

            # Mostrar tabla en la aplicación
            st.markdown("### Tabla consolidada:")
            st.dataframe(tabla_df)

            # Crear gráficos de barras por tipo de documento
            st.markdown("### Gráficos de barras: porcentaje relativo del valor por tipo de documento")

            # Lista de gráficos que se mostrarán
            all_figures = []
            for tipo_doc in tipo_documentos:
                fig, axes = plt.subplots(1, 2, figsize=(16, 6), sharey=True)
                fig.suptitle(f"Porcentaje relativo de {tipo_doc}", fontsize=16)

                for ax, grado in zip(axes, grados):
                    # Filtrar los datos para el gráfico
                    df_filtro = tabla_df[(tabla_df["Tipo Doc"] == tipo_doc) & (tabla_df["Grado"] == grado)]
                    total_anual = df_filtro["Total Anual"].values[0]

                    # Omitir gráficos si el total anual es 0
                    if total_anual == 0:
                        continue

                    # Calcular el porcentaje de cada mes respecto al total anual
                    porcentajes = (df_filtro[meses_orden].values.flatten() / total_anual) * 100

                    # Crear el gráfico de barras
                    ax.bar(meses_orden, porcentajes, color='skyblue', width=0.6)
                    ax.set_title(grado, fontsize=14)
                    ax.set_xlabel("Mes", fontsize=12)
                    ax.set_ylabel("Porcentaje (%)", fontsize=12)
                    ax.set_ylim(0, 100)
                    ax.set_xticks(range(len(meses_orden)))
                    ax.set_xticklabels(meses_orden, rotation=45, ha='right')

                    # Agregar etiquetas a las barras
                    for i, porcentaje in enumerate(porcentajes):
                        ax.text(i, porcentaje + 1, f"{porcentaje:.1f}%", ha='center', va='bottom', fontsize=10)

                # Mostrar el gráfico en la aplicación
                st.pyplot(fig)

                # Añadir el gráfico a la lista de figuras para el PDF
                all_figures.append(fig)

            # Crear un PDF para guardar los gráficos
            def crear_pdf(figures):
                pdf_output = BytesIO()
                with PdfPages(pdf_output) as pdf:
                    for fig in figures:
                        pdf.savefig(fig)
                        plt.close(fig)
                return pdf_output.getvalue()

            # Botón para descargar el PDF de los gráficos
            st.markdown("### Descargar gráficos en PDF")
            pdf_data = crear_pdf(all_figures)
            st.download_button(
                label="Descargar gráficos en PDF",
                data=pdf_data,
                file_name="graficos_dian.pdf",
                mime="application/pdf"
            )

            # Generar archivo Excel para descargar
            @st.cache_data
            def convertir_a_excel(dataframe):
                output = BytesIO()
                wb = Workbook()
                ws = wb.active
                ws.title = "Resultados"

                for r_idx, row in enumerate(dataframe_to_rows(dataframe, index=False, header=True), 1):
                    ws.append(row)

                # Agregar gráficos al Excel
                for tipo_doc in tipo_documentos:
                    for grado in grados:
                        df_filtro = tabla_df[(tabla_df["Tipo Doc"] == tipo_doc) & (tabla_df["Grado"] == grado)]
                        total_anual = df_filtro["Total Anual"].values[0]

                        if total_anual == 0:
                            continue

                        porcentajes = (df_filtro[meses_orden].values.flatten() / total_anual) * 100
                        fig, ax = plt.subplots(figsize=(8, 4))
                        ax.bar(meses_orden, porcentajes, color='skyblue', width=0.6)
                        ax.set_title(f"{tipo_doc} - {grado}")
                        ax.set_xlabel("Mes")
                        ax.set_ylabel("Porcentaje (%)")

                        for i, porcentaje in enumerate(porcentajes):
                            ax.text(i, porcentaje + 1, f"{porcentaje:.1f}%", ha='center', va='bottom', fontsize=8)

                        image_stream = BytesIO()
                        plt.savefig(image_stream, format='png')
                        plt.close(fig)

                        image_stream.seek(0)
                        img = Image(image_stream)
                        ws.add_image(img, f"B{len(ws['A']) + 5}")

                wb.save(output)
                return output.getvalue()

            # Archivo para descargar
            excel_data = convertir_a_excel(tabla_df)
            st.download_button(
                label="Descargar tabla consolidada en Excel",
                data=excel_data,
                file_name="analisis_dian.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error procesando el archivo: {e}")
else:
    st.write("Por favor, sube un archivo Excel para comenzar el análisis.")
