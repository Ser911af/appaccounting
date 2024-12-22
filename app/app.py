import streamlit as st
import time
from groq import Groq
from typing import Generator
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import pandas as pd

# Título de la aplicación
st.title("Groq Sergio Bot")

# Declaramos el cliente de Groq
client = Groq(
    api_key=st.secrets["ngroqAPIKey"],  # Cargamos la API key del .streamlit/secrets.toml
)

# Lista de modelos para elegir
modelos = ['llama3-8b-8192', 'llama3-70b-8192', 'mixtral-8x7b-32768']

def generate_chat_responses(chat_completion) -> Generator[str, None, None]:
    """ Genera respuestas de chat mostrando caracter por caracter. """
    for chunk in chat_completion:
        if chunk.choices[0].delta.content:
            yield chunk.choices[0].delta.content

# Función para generar un PDF con el contenido del informe
def create_pdf(report_content):
    pdf_filename = "/tmp/Informe_Analisis.pdf"
    c = canvas.Canvas(pdf_filename, pagesize=letter)
    c.drawString(100, 750, report_content)
    c.save()
    return pdf_filename

# Inicializamos el historial de chat
if "messages" not in st.session_state:
    st.session_state.messages = []

# Muestra mensajes de chat desde la historia en la aplicación
with st.container():
    for message in st.session_state.messages:
        with st.chat_message(message["role"]):
            st.markdown(message["content"])

# Mostramos la lista de modelos en el sidebar
parModelo = st.sidebar.selectbox('Modelos', options=modelos, index=0)

# Botón para cargar los datos de análisis
uploaded_file = st.file_uploader("Sube tu archivo Excel", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # Procesar el DataFrame, crear la tabla y generar un resumen de la tabla
    required_columns = ["Fecha Emisión", "Total", "IVA", "Tipo de documento", "Grupo"]
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        st.error(f"El archivo no contiene las columnas requeridas: {', '.join(missing_columns)}")
    else:
        df["Fecha Emisión"] = pd.to_datetime(df["Fecha Emisión"], format='%d-%m-%Y', errors="coerce")
        df["Total"] = pd.to_numeric(df["Total"], errors='coerce')
        df["IVA"] = pd.to_numeric(df["IVA"], errors='coerce')
        df["Base"] = (df["Total"].fillna(0) - df["IVA"].fillna(0)).round(0)

        # Crear la tabla consolidada (igual que antes)
        month_mapping = {
            1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
            7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
        }
        df["Mes"] = df["Fecha Emisión"].dt.month.map(month_mapping)
        meses_orden = list(month_mapping.values())

        # Crear la tabla consolidada
        tipo_documentos = df["Tipo de documento"].unique()
        grados = ["Emitido", "Recibido"]
        tabla_resultados = []

        for tipo_doc in tipo_documentos:
            for grado in grados:
                df_filtro = df[(df["Tipo de documento"] == tipo_doc) & (df["Grupo"] == grado)]
                suma_por_mes = df_filtro.groupby("Mes")["Base"].sum().reindex(meses_orden, fill_value=0)
                total_anual = suma_por_mes.sum()
                fila = [tipo_doc, grado] + list(suma_por_mes.values) + [total_anual]
                tabla_resultados.append(fila)

        tabla_df = pd.DataFrame(tabla_resultados, columns=["Tipo Doc", "Grado"] + meses_orden + ["Total Anual"])
        tabla_resumen = tabla_df.describe().to_string()

        st.dataframe(tabla_df)
        st.markdown("### Resumen de la tabla:")
        st.text(tabla_resumen)

        # Botón para enviar la solicitud al modelo
        prompt = f"""
        Genera un informe detallado basado en los siguientes datos de facturación:
        {tabla_resumen}
        El informe debe incluir:
        - Análisis de los meses con mayores y menores ingresos.
        - Observaciones sobre tendencias.
        - Recomendaciones basadas en los datos presentados.
        """

        # Generar respuesta del modelo
        if prompt:
            st.chat_message("user").markdown(prompt)
            st.session_state.messages.append({"role": "user", "content": prompt})
            try:
                with st.spinner('Generando informe...'):
                    chat_completion = client.chat.completions.create(
                        model=parModelo,
                        messages=[{"role": m["role"], "content": m["content"]} for m in st.session_state.messages],
                        stream=True
                    )
                    chat_responses_generator = generate_chat_responses(chat_completion)
                    full_response = ''.join(list(chat_responses_generator))
                    st.chat_message("assistant").markdown(full_response)
                    st.session_state.messages.append({"role": "assistant", "content": full_response})

                    # Crear y descargar el PDF
                    pdf_filename = create_pdf(full_response)
                    st.download_button(
                        label="Descargar Informe en PDF",
                        data=open(pdf_filename, "rb").read(),
                        file_name="Informe_Analisis.pdf",
                        mime="application/pdf"
                    )

            except Exception as e:
                st.error(e)
