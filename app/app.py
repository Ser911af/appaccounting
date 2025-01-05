import pandas as pd
import streamlit as st

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

            # Crear dos columnas 'Base' con diferentes fórmulas
            df["Base_Total"] = (df["Total"].fillna(0) - df["IVA"].fillna(0)).round(0)
            df["Base_IVA"] = (df["IVA"].fillna(0)).round(0)

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

            # Función para crear tablas consolidadas
            def crear_tabla_consolidada(base_columna):
                tabla_resultados = []
                for tipo_doc in tipo_documentos:
                    for grado in grados:
                        # Filtrar datos por 'Tipo de documento' y 'Grupo'
                        df_filtro = df[(df["Tipo de documento"] == tipo_doc) & (df["Grupo"] == grado)]

                        # Sumar la columna Base por mes
                        suma_por_mes = (
                            df_filtro.groupby("Mes")[base_columna].sum()
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
                return tabla_df

            # Opciones para el usuario
            opcion = st.radio("Selecciona el tipo de análisis que deseas realizar:", 
                              ("Base Total (Total - IVA)", "Base IVA (IVA solamente)"))

            # Crear las tablas basadas en la selección del usuario
            if opcion == "Base Total (Total - IVA)":
                tabla_df = crear_tabla_consolidada("Base_Total")
                st.markdown("### Tabla consolidada basada en 'Base Total (Total - IVA)':")
                st.dataframe(tabla_df)

            elif opcion == "Base IVA (IVA solamente)":
                tabla_df = crear_tabla_consolidada("Base_IVA")
                st.markdown("### Tabla consolidada basada en 'Base IVA (IVA solamente)':")
                st.dataframe(tabla_df)

    except Exception as e:
        st.error(f"Error procesando el archivo: {e}")
else:
    st.write("Por favor, sube un archivo Excel para comenzar.")
