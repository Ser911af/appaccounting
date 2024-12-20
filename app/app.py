import pandas as pd
import streamlit as st
from io import BytesIO

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
            if df["Fecha Emisión"].isnull().any():
                st.warning("Algunas fechas no pudieron ser convertidas correctamente.")

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
                    total = suma_por_mes.sum()

                    # Crear fila de resultados
                    fila = [tipo_doc, grado] + list(suma_por_mes.values) + [total]
                    tabla_resultados.append(fila)

            # Crear DataFrame con la tabla consolidada
            columnas = ["Tipo Doc", "Grado"] + meses_orden + ["Total"]
            tabla_df = pd.DataFrame(tabla_resultados, columns=columnas)

            # Redondear valores a enteros
            tabla_df = tabla_df.round(0)

            # Mostrar tabla en la aplicación
            st.markdown("### Tabla consolidada:")
            st.dataframe(tabla_df)

            # Generar archivo Excel para descargar
            @st.cache_data
            def convertir_a_excel(dataframe):
                output = BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    dataframe.to_excel(writer, index=False, sheet_name="Resultados")
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
