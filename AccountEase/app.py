import pandas as pd
import streamlit as st

# Título de la aplicación
st.title("DIAN Report Analyzer")
st.subheader("Carga tu reporte DIAN y obtén un análisis básico del archivo")

# Subida del archivo
uploaded_file = st.file_uploader("Sube tu archivo Excel", type=["xlsx"])

if uploaded_file:
    # Leer el archivo Excel
    df = pd.read_excel(uploaded_file)
    
    # Mostrar una vista previa del archivo
    st.write("Vista previa de los datos cargados:")
    st.dataframe(df.head())

    # Generar el reporte básico
    st.markdown("### Resumen del archivo:")
    
    # Información sobre las columnas
    st.write(f"**Número de columnas:** {df.shape[1]}")
    st.write(f"**Nombres de columnas:** {', '.join(df.columns)}")
    
    # Información sobre los registros
    st.write(f"**Número total de registros:** {df.shape[0]}")

    # Información adicional por columna
    st.markdown("### Detalles de las columnas:")
    columnas_info = {
        "Nombre": df.columns,
        "Tipo de dato": [df[col].dtype for col in df.columns],
        "Registros únicos": [df[col].nunique() for col in df.columns],
        "Valores nulos": [df[col].isnull().sum() for col in df.columns],
    }
    columnas_df = pd.DataFrame(columnas_info)
    st.dataframe(columnas_df)

    # Opcional: Descargar el reporte como Excel
    @st.cache
    def convert_df_to_excel(dataframe):
        return dataframe.to_excel(index=False, engine="openpyxl")
    
    excel_report = convert_df_to_excel(columnas_df)
    st.download_button(
        "Descargar reporte de columnas",
        data=excel_report,
        file_name="reporte_columnas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.write("Por favor, sube un archivo Excel para comenzar el análisis.")
