import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from matplotlib.backends.backend_pdf import PdfPages

# --- UTILIDADES ---
def convertir_a_excel(dataframe):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        dataframe.to_excel(writer, index=False, sheet_name='Consolidado')
    return output.getvalue()

def convertir_a_pdf(figs):
    output = BytesIO()
    with PdfPages(output) as pdf:
        for fig in figs:
            pdf.savefig(fig)
    output.seek(0)
    return output

def detectar_hoja_y_leer(excel_file):
    xls = pd.ExcelFile(excel_file)
    sheet = xls.sheet_names[0]  # Primera hoja por defecto
    for skip in range(10):  # Probar con primeras 10 filas
        df = pd.read_excel(xls, sheet_name=sheet, skiprows=skip)
        if {'Fecha Emisión', 'Total', 'IVA', 'Tipo de documento', 'Grupo'}.issubset(df.columns):
            return df
    raise ValueError("No se encontraron columnas requeridas en las primeras filas")

def preparar_dataframe(df):
    df["Fecha Emisión"] = pd.to_datetime(df["Fecha Emisión"], format='%d-%m-%Y', errors="coerce")
    df["Total"] = pd.to_numeric(df["Total"], errors='coerce')
    df["IVA"] = pd.to_numeric(df["IVA"], errors='coerce')
    df["Base_Total"] = (df["Total"].fillna(0) - df["IVA"].fillna(0)).round(0)
    df["Base_IVA"] = df["IVA"].fillna(0).round(0)
    df = df.dropna(subset=["Fecha Emisión"])
    month_mapping = {
        1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
        5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
        9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
    }
    df["Mes"] = df["Fecha Emisión"].dt.month.map(month_mapping)
    return df

def generar_tabla(df, base_columna):
    meses_presentes = df["Mes"].dropna().unique().tolist()
    meses_presentes = sorted(meses_presentes, key=lambda x: list(month_mapping.values()).index(x))
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
    return pd.DataFrame(tabla_resultados, columns=columnas).round(0), meses_presentes

def resumen_por_grupo(df, base_columna):
    resumen = df.groupby("Grupo")[base_columna].sum().reset_index()
    resumen["%"] = (resumen[base_columna] / resumen[base_columna].sum()) * 100
    return resumen.round(0)

def graficar_total_mensual(df, base_columna, meses):
    total_por_mes = df.groupby("Mes")[base_columna].sum().reindex(meses, fill_value=0)
    total_por_mes_millones = total_por_mes // 1_000_000
    total_anual = total_por_mes.sum()
    porcentajes = (total_por_mes / total_anual) * 100
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.plot(meses, total_por_mes_millones, marker='o', color='b', linestyle='-')
    ax.set_title("Evolución mensual (en millones de pesos)", fontsize=16)
    ax.set_xlabel("Mes")
    ax.set_ylabel("Total (Millones de Pesos)")
    for i, value in enumerate(total_por_mes_millones):
        ax.text(i, value, f"{int(value):,}M\n({int(porcentajes[i])}%)".replace(",", "."), ha='center')
    return fig

def graficar_barras_por_tipo(tabla_df, meses):
    figs = []
    tipo_documentos = tabla_df["Tipo Doc"].unique()
    for tipo_doc in tipo_documentos:
        fig, axes = plt.subplots(1, 2, figsize=(16, 6), sharey=True)
        fig.suptitle(f"Porcentaje relativo de {tipo_doc}", fontsize=16)
        for ax, grado in zip(axes, ["Emitido", "Recibido"]):
            df_filtro = tabla_df[(tabla_df["Tipo Doc"] == tipo_doc) & (tabla_df["Grado"] == grado)]
            if not df_filtro.empty:
                porcentajes = (df_filtro[meses].values.flatten() / df_filtro["Total Anual"].values[0]) * 100
                ax.bar(meses, porcentajes, color='skyblue', width=0.6)
                ax.set_title(grado)
                for i, porcentaje in enumerate(porcentajes):
                    ax.text(i, porcentaje + 1, f"{porcentaje:.0f}%", ha='center')
        figs.append(fig)
    return figs

# --- INTERFAZ STREAMLIT ---
st.title("DIAN Report Analyzer")
st.subheader("Carga tu reporte DIAN y obtén un análisis detallado del archivo")

uploaded_file = st.file_uploader("Sube tu archivo Excel", type=["xlsx"])

if uploaded_file:
    try:
        df = detectar_hoja_y_leer(uploaded_file)

        required_columns = ["Fecha Emisión", "Total", "IVA", "Tipo de documento", "Grupo"]
        if not all(col in df.columns for col in required_columns):
            st.error("Faltan columnas requeridas en el archivo")
        else:
            df = preparar_dataframe(df)

            analisis = st.radio("Selecciona el tipo de análisis", ["Base con Total e IVA", "Base con solo IVA"])
            base_columna = "Base_Total" if analisis == "Base con Total e IVA" else "Base_IVA"

            tabla_df, meses_presentes = generar_tabla(df, base_columna)

            with st.expander("Ver tabla consolidada"):
                st.dataframe(tabla_df)

            resumen = resumen_por_grupo(df, base_columna)
            st.markdown("### Resumen por grupo")
            st.dataframe(resumen)

            st.markdown("### Descargar tabla en Excel")
            excel_data = convertir_a_excel(tabla_df)
            st.download_button("Descargar Excel", data=excel_data, file_name="tabla_consolidada.xlsx")

            st.markdown("### Gráficos")
            fig_linea = graficar_total_mensual(df, base_columna, meses_presentes)
            st.pyplot(fig_linea)

            figs_barras = graficar_barras_por_tipo(tabla_df, meses_presentes)
            for fig in figs_barras:
                st.pyplot(fig)

            st.markdown("### Exportar gráficos a PDF")
            if st.button("Descargar PDF"):
                pdf_data = convertir_a_pdf([fig_linea] + figs_barras)
                st.download_button("Descargar PDF", data=pdf_data, file_name="graficos.pdf", mime="application/pdf")

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")
