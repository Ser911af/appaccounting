import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO

# Función para convertir el DataFrame en un archivo Excel descargable
def convertir_a_excel(dataframe):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        dataframe.to_excel(writer, index=False, sheet_name='Consolidado')
    processed_data = output.getvalue()
    return processed_data

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

            # Crear dos columnas 'Base' con diferentes fórmulas
            df["Base_Total"] = (df["Total"].fillna(0) - df["IVA"].fillna(0)).round(0)
            df["Base_IVA"] = (df["IVA"].fillna(0)).round(0)

            # Extraer el nombre del mes
            month_mapping = {
                1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
                5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
                9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
            }
            df["Mes"] = df["Fecha Emisión"].dt.month.map(month_mapping)

            # Ordenar meses correctamente
            meses_orden = list(month_mapping.values())
            df["Mes"] = pd.Categorical(df["Mes"], categories=meses_orden, ordered=True)

            # Obtener valores únicos de 'Tipo de documento' y 'Grupo'
            tipo_documentos = df["Tipo de documento"].unique()
            grados = ["Emitido", "Recibido"]

            # Opción para seleccionar el tipo de análisis
            analisis = st.selectbox(
                "Selecciona el tipo de análisis",
                ["Base con Total e IVA", "Base con solo IVA"]
            )

            # Crear tabla consolidada basada en el análisis seleccionado
            tabla_resultados = []
            for tipo_doc in tipo_documentos:
                for grado in grados:
                    # Filtrar datos por 'Tipo de documento' y 'Grupo'
                    df_filtro = df[(df["Tipo de documento"] == tipo_doc) & (df["Grupo"] == grado)]

                    # Sumar bases por mes según el análisis seleccionado
                    base_col = "Base_Total" if analisis == "Base con Total e IVA" else "Base_IVA"
                    suma_por_mes = (
                        df_filtro.groupby("Mes")[base_col].sum()
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
            st.markdown(f"### Tabla consolidada: {analisis}")
            st.dataframe(tabla_df)

            # Descargar tabla consolidada en Excel
            excel_data = convertir_a_excel(tabla_df)
            st.download_button(
                label="Descargar tabla en Excel",
                data=excel_data,
                file_name="tabla_consolidada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # Visualización gráfica
            st.markdown("### Visualización gráfica")

            # Gráfico principal (líneas)
            base_col = "Base_Total" if analisis == "Base con Total e IVA" else "Base_IVA"
            suma_mensual_total = df.groupby("Mes")[base_col].sum().reindex(meses_orden, fill_value=0)
            fig, ax = plt.subplots(figsize=(12, 6))
            suma_mensual_total.plot(kind="line", ax=ax, marker="o", color="blue")
            ax.set_title("Total Mensual Consolidado", fontsize=16)
            ax.set_ylabel("Total", fontsize=12)
            ax.set_xlabel("Mes", fontsize=12)
            ax.grid(True)
            st.pyplot(fig)

            # Gráficos secundarios (barras)
            for tipo_doc in tipo_documentos:
                st.markdown(f"#### {tipo_doc}")
                col1, col2 = st.columns(2)
                for i, grado in enumerate(grados):
                    df_filtro = df[(df["Tipo de documento"] == tipo_doc) & (df["Grupo"] == grado)]
                    suma_por_mes = (
                        df_filtro.groupby("Mes")[base_col].sum()
                        .reindex(meses_orden, fill_value=0)
                    )
                    fig, ax = plt.subplots(figsize=(5, 3))
                    suma_por_mes.plot(kind="bar", ax=ax, color="skyblue", edgecolor="black")
                    ax.set_title(f"{grado}", fontsize=12)
                    ax.set_ylabel("Total", fontsize=10)
                    ax.set_xlabel("Mes", fontsize=10)
                    ax.set_xticklabels(meses_orden, rotation=45)
                    if i == 0:
                        col1.pyplot(fig)
                    else:
                        col2.pyplot(fig)

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")
