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

# Función para convertir el DataFrame en un archivo Excel descargable
def convertir_a_excel(dataframe):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        dataframe.to_excel(writer, index=False, sheet_name='Consolidado')
    return output.getvalue()

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

            # Crear columnas de bases
            df["Base_Total"] = (df["Total"].fillna(0) - df["IVA"].fillna(0)).round(0)
            df["Base_IVA"] = df["IVA"].fillna(0).round(0)

            # Extraer mes en texto
            month_mapping = {
                1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
                5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
                9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
            }
            df["Mes"] = df["Fecha Emisión"].dt.month.map(month_mapping)
            meses_orden = list(month_mapping.values())
            df["Mes"] = pd.Categorical(df["Mes"], categories=meses_orden, ordered=True)

            # Opción para seleccionar el tipo de análisis
            analisis = st.radio(
                "Selecciona el tipo de análisis",
                ["Base con Total e IVA", "Base con solo IVA"]
            )

            # Configurar columna base según análisis
            base_columna = "Base_Total" if analisis == "Base con Total e IVA" else "Base_IVA"

            # Calcular tabla consolidada
            tipo_documentos = df["Tipo de documento"].unique()
            grados = ["Emitido", "Recibido"]
            tabla_resultados = []
            for tipo_doc in tipo_documentos:
                for grado in grados:
                    df_filtro = df[(df["Tipo de documento"] == tipo_doc) & (df["Grupo"] == grado)]
                    suma_por_mes = df_filtro.groupby("Mes")[base_columna].sum().reindex(meses_orden, fill_value=0)
                    total_anual = suma_por_mes.sum()
                    fila = [tipo_doc, grado] + list(suma_por_mes.values) + [total_anual]
                    tabla_resultados.append(fila)

            columnas = ["Tipo Doc", "Grado"] + meses_orden + ["Total Anual"]
            tabla_df = pd.DataFrame(tabla_resultados, columns=columnas).round(0)

            # Mostrar tabla consolidada
            st.markdown("### Tabla consolidada:")
            st.dataframe(tabla_df)

            # Botón para descargar el archivo Excel
            st.markdown("### Descargar tabla en Excel")
            excel_data = convertir_a_excel(tabla_df)
            st.download_button(
                label="Descargar tabla en Excel",
                data=excel_data,
                file_name="tabla_consolidada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # Gráfico principal: Línea del total mensual
            st.markdown("### Gráfico: Evolución del total mensual")
            total_por_mes = df.groupby("Mes")[base_columna].sum().reindex(meses_orden, fill_value=0)
            total_por_mes_millones = total_por_mes // 1_000_000

            fig_linea, ax_linea = plt.subplots(figsize=(10, 6))
            ax_linea.plot(total_por_mes.index, total_por_mes_millones, marker='o', color='b', linestyle='-')
            ax_linea.set_title("Evolución mensual (en millones de pesos)", fontsize=16)
            ax_linea.set_xlabel("Mes", fontsize=12)
            ax_linea.set_ylabel("Total (Millones de Pesos)", fontsize=12)
            ax_linea.grid(True, linestyle='--', alpha=0.6)
            ax_linea.set_xticks(range(len(meses_orden)))
            ax_linea.set_xticklabels(meses_orden, rotation=45)

            total_anual = total_por_mes.sum()
            porcentajes = (total_por_mes / total_anual) * 100
            for i, value in enumerate(total_por_mes_millones):
                ax_linea.text(
                    i, value,
                    f"{int(value):,}M\n({int(porcentajes[i])}%)".replace(",", "."),
                    ha='center', va='bottom', fontsize=10
                )

            ax_linea.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f"{int(x):,}".replace(",", ".")))
            st.pyplot(fig_linea)

            # Gráficos de barras por tipo de documento
            st.markdown("### Gráficos de barras por tipo de documento")
            for tipo_doc in tipo_documentos:
                fig, axes = plt.subplots(1, 2, figsize=(16, 6), sharey=True)
                fig.suptitle(f"Porcentaje relativo de {tipo_doc}", fontsize=16)

                for ax, grado in zip(axes, grados):
                    df_filtro = tabla_df[(tabla_df["Tipo Doc"] == tipo_doc) & (tabla_df["Grado"] == grado)]
                    if not df_filtro.empty:
                        porcentajes = (df_filtro[meses_orden].values.flatten() / df_filtro["Total Anual"].values[0]) * 100
                        ax.bar(meses_orden, porcentajes, color='skyblue', width=0.6)
                        ax.set_title(grado, fontsize=14)
                        ax.set_xlabel("Mes", fontsize=12)
                        ax.set_ylabel("Porcentaje (%)", fontsize=12)
                        ax.set_ylim(0, 100)
                        ax.set_xticks(range(len(meses_orden)))
                        ax.set_xticklabels(meses_orden, rotation=45)

                        for i, porcentaje in enumerate(porcentajes):
                            ax.text(i, porcentaje + 1, f"{porcentaje:.0f}%", ha='center', va='bottom', fontsize=10)

                st.pyplot(fig)

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")
