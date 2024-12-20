import pandas as pd
import streamlit as st
from io import BytesIO

# Título de la aplicación
st.title("DIAN Report Analyzer")
st.subheader("Carga tu reporte DIAN y obtén un análisis detallado del archivo")

# Subida del archivo
uploaded_file = st.file_uploader("Sube tu archivo Excel", type=["xlsx"])

if uploaded_file:
    # Leer el archivo Excel
    df = pd.read_excel(uploaded_file)

    # Crear nueva variable 'Base'
    df["Base"] = df["Total"] - df["IVA"]

    # Agregar columna con nombre del mes basado en 'Fecha Emisión'
    df["Mes"] = pd.to_datetime(df["Fecha Emisión"]).dt.month_name()

    # Ordenar meses correctamente
    meses_orden = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ]
    df["Mes"] = pd.Categorical(df["Mes"], categories=meses_orden, ordered=True)

    # Obtener valores únicos de 'Tipo de documento' y 'Grado'
    tipo_documentos = df["Tipo de documento"].unique()
    grados = ["Emitido", "Recibido"]

    # Crear tabla consolidada
    tabla_resultados = []

    for tipo_doc in tipo_documentos:
        for grado in grados:
            # Filtrar datos por 'Tipo de documento' y 'Grado'
            df_filtro = df[(df["Tipo de documento"] == tipo_doc) & (df["Grupo"] == grado)]

            # Sumar 'Base' por mes
            suma_por_mes = (
                df_filtro.groupby("Mes")["Base"].sum()
                .reindex(meses_orden, fill_value=0)  # Asegurar que todos los meses aparezcan
            )

            # Calcular total anual
            total = suma_por_mes.sum()

            # Crear fila de resultados
            fila = [tipo_doc, grado] + list(suma_por_mes.values) + [total]
            tabla_resultados.append(fila)

    # Crear DataFrame con la tabla consolidada
    columnas = ["Tipo Doc", "Grado"] + meses_orden + ["Total"]
    tabla_df = pd.DataFrame(tabla_resultados, columns=columnas)

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
else:
    st.write("Por favor, sube un archivo Excel para comenzar el análisis.")
