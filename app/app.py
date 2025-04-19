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

# --- FUNCIONES AUXILIARES ---

def convertir_a_excel(dataframe):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        dataframe.to_excel(writer, index=False, sheet_name='Consolidado')
    return output.getvalue()

def obtener_meses_presentes(df, columna_mes="Mes"):
    meses_orden = [
        "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
        "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
    ]
    meses_presentes = [mes for mes in meses_orden if mes in df[columna_mes].unique()]
    return meses_presentes

def graficar_linea_evolucion(df, columna_valor, titulo="Evolución mensual"):
    meses = obtener_meses_presentes(df)
    total_por_mes = df.groupby("Mes")[columna_valor].sum().reindex(meses, fill_value=0)
    total_por_mes_millones = total_por_mes // 1_000_000
    total_anual = total_por_mes.sum()
    porcentajes = (total_por_mes / total_anual) * 100

    fig, ax = plt.subplots(figsize=(10, 6))
    ax.plot(meses, total_por_mes_millones, marker='o', color='b', linestyle='-')
    ax.set_title(titulo, fontsize=16)
    ax.set_xlabel("Mes", fontsize=12)
    ax.set_ylabel("Total (Millones de Pesos)", fontsize=12)
    ax.grid(True, linestyle='--', alpha=0.6)
    ax.set_xticks(range(len(meses)))
    ax.set_xticklabels(meses, rotation=45)

    for i, value in enumerate(total_por_mes_millones):
        ax.text(i, value, f"{int(value):,}M\n({int(porcentajes[i])}%)".replace(",", "."),
                ha='center', va='bottom', fontsize=10)

    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f"{int(x):,}".replace(",", ".")))
    st.pyplot(fig)

def graficar_barras_tipo_documento(df_tabla, tipo_doc, meses):
    grados = ["Emitido", "Recibido"]
    fig, axes = plt.subplots(1, 2, figsize=(16, 6), sharey=True)
    fig.suptitle(f"Porcentaje relativo de {tipo_doc}", fontsize=16)

    for ax, grado in zip(axes, grados):
        df_filtro = df_tabla[(df_tabla["Tipo Doc"] == tipo_doc) & (df_tabla["Grado"] == grado)]
        if not df_filtro.empty:
            porcentajes = (df_filtro[meses].values.flatten() / df_filtro["Total Anual"].values[0]) * 100
            ax.bar(meses, porcentajes, color='skyblue', width=0.6)
            ax.set_title(grado, fontsize=14)
            ax.set_xlabel("Mes", fontsize=12)
            ax.set_ylabel("Porcentaje (%)", fontsize=12)
            ax.set_ylim(0, 100)
            ax.set_xticks(range(len(meses)))
            ax.set_xticklabels(meses, rotation=45)

            for i, porcentaje in enumerate(porcentajes):
                ax.text(i, porcentaje + 1, f"{porcentaje:.0f}%", ha='center', va='bottom', fontsize=10)

    st.pyplot(fig)

# --- PROCESAMIENTO PRINCIPAL ---
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
            df["Mes"] = pd.Categorical(df["Mes"], categories=list(month_mapping.values()), ordered=True)

            meses_presentes = obtener_meses_presentes(df)

            # Selección del análisis
            analisis = st.radio("Selecciona el tipo de análisis", ["Base con Total e IVA", "Base con solo IVA"])
            base_columna = "Base_Total" if analisis == "Base con Total e IVA" else "Base_IVA"

            # Construcción de la tabla consolidada
            tipo_documentos = df["Tipo de documento"].unique()
            grados = ["Emitido", "Recibido"]
            tabla_resultados = []

            for tipo_doc in tipo_documentos:
                for grado in grados:
                    df_filtro = df[(df["Tipo de documento"] == tipo_doc) & (df["Grupo"] == grado)]
                    suma_por_mes = df_filtro.groupby("Mes")[base_columna].sum().reindex(meses_presentes, fill_value=0)
                    total_anual = suma_por_mes.sum()
                    fila = [tipo_doc, grado] + list(suma_por_mes.values) + [total_anual]
                    tabla_resultados.append(fila)

            columnas = ["Tipo Doc", "Grado"] + meses_presentes + ["Total Anual"]
            tabla_df = pd.DataFrame(tabla_resultados, columns=columnas).round(0)

            # Mostrar tabla consolidada
            st.markdown("### Tabla consolidada:")
            st.dataframe(tabla_df)

            # Descargar Excel
            st.markdown("### Descargar tabla en Excel")
            excel_data = convertir_a_excel(tabla_df)
            st.download_button(
                label="Descargar tabla en Excel",
                data=excel_data,
                file_name="tabla_consolidada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # Gráfico de evolución mensual
            st.markdown("### Gráfico: Evolución del total mensual")
            graficar_linea_evolucion(df, base_columna)

            # Gráficos de barras por tipo de documento
            st.markdown("### Gráficos de barras por tipo de documento")
            for tipo_doc in tipo_documentos:
                graficar_barras_tipo_documento(tabla_df, tipo_doc, meses_presentes)

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")
