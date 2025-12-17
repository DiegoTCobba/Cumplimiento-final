import streamlit as st
import pandas as pd
import requests
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from datetime import datetime

# ===============================
# CONFIGURACIÓN GENERAL
# ===============================
st.set_page_config(
    page_title="Cumplimiento – Rechazos BCP",
    layout="wide"
)

API_URL = "https://q6capnpv09.execute-api.us-east-1.amazonaws.com/OPS/kpayout/v1/payout_process/reject_invoices_batch"

HEADERS = {
    # "Authorization": "Bearer TU_TOKEN"
}

CODIGO_RECHAZO = "R016"
DESCRIPCION_RECHAZO = "CUENTA INVÁLIDA"

# ===============================
# FUNCIONES
# ===============================
def generar_excel_rechazo(referencias):
    df = pd.DataFrame({
        "Referencia": referencias,
        "Estado": ["Rechazada"] * len(referencias),
        "Codigo de Rechazo": [CODIGO_RECHAZO] * len(referencias),
        "Descripcion de Rechazo": [DESCRIPCION_RECHAZO] * len(referencias)
    })

    buffer = BytesIO()
    df.to_excel(buffer, index=False, engine="openpyxl")
    buffer.seek(0)

    wb = load_workbook(buffer)
    ws = wb.active

    for col in ws.iter_cols(min_col=1, max_col=1, min_row=2):
        for cell in col:
            cell.number_format = "@"

    buffer_final = BytesIO()
    wb.save(buffer_final)
    buffer_final.seek(0)

    return buffer_final


def enviar_rechazo_api(buffer_excel):
    files = {
        "edt": (
            "RechazoBCP.xlsx",
            buffer_excel,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    }

    response = requests.post(
        API_URL,
        headers=HEADERS,
        files=files,
        timeout=60
    )

    return response


def generar_excel_respaldo(df):
    fecha = datetime.now().strftime("%d.%m.%y")
    buffer = BytesIO()

    df_export = df.drop(columns=["Seleccionar"], errors="ignore")

    df_export.to_excel(
        buffer,
        index=False,
        engine="openpyxl",
        sheet_name="Registros Observados"
    )

    buffer.seek(0)
    return buffer, f"Registros_Observados_{fecha}.xlsx"


def generar_formato_due_diligence(df):
    """
    Genera el Formato Due Diligence usando la PLANTILLA OFICIAL
    """
    fecha_excel = datetime.now().strftime("%d/%m/%Y")
    fecha_archivo = datetime.now().strftime("%d.%m.%y")

    # Cargar plantilla EXACTA
    wb = load_workbook("plantillas/Formato_Due_Diligence_Template.xlsx")
    ws = wb.active  # o wb["Due Diligence"]

    # Escribir fecha (celda según tu plantilla)
    ws["D10"] = fecha_excel

    # Fila donde empiezan los registros (según tu imagen)
    fila_inicio = 13

    # Limpiar registros anteriores
    for row in ws.iter_rows(min_row=fila_inicio, max_col=5):
        for cell in row:
            cell.value = None

    # Insertar clientes observados (>30K)
    for i, row in enumerate(df.itertuples(), start=0):
        ws[f"C{fila_inicio + i}"] = row.DOCUMENTO
        ws[f"D{fila_inicio + i}"] = str(row.NUMERO_DOCUMENTO)
        ws[f"E{fila_inicio + i}"] = row.NOMBRE

    # Guardar en memoria
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    return buffer, f"Formato Due Dilligence {fecha_archivo}.xlsx"

# ===============================
# INTERFAZ STREAMLIT
# ===============================
st.title("🚨 Cumplimiento – Rechazo de Clientes (>30K)")
st.write("Carga archivos Excel, revisa clientes observados y genera evidencias oficiales.")

uploaded_files = st.file_uploader(
    "📂 Cargar uno o más archivos Excel",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files:
    dataframes = []

    for file in uploaded_files:
        try:
            df = pd.read_excel(file)

            columnas = df.iloc[:, [1, 2, 3, 8, 12]].copy()
            columnas.columns = [
                "DOCUMENTO",
                "NUMERO_DOCUMENTO",
                "NOMBRE",
                "REFERENCIA",
                "MONTO"
            ]

            columnas["MONTO"] = pd.to_numeric(columnas["MONTO"], errors="coerce")
            columnas["REFERENCIA"] = columnas["REFERENCIA"].astype(str)

            filtrado = columnas[columnas["MONTO"] > 30000].copy()
            filtrado["Archivo_Origen"] = file.name
            filtrado["Seleccionar"] = True

            dataframes.append(filtrado)

        except Exception as e:
            st.error(f"Error procesando {file.name}: {e}")

    if dataframes:
        resultado_final = pd.concat(dataframes, ignore_index=True)

        st.subheader("📋 Registros detectados (>30K)")
        editable_df = st.data_editor(
            resultado_final,
            hide_index=True,
            use_container_width=True
        )

        seleccionados = editable_df[editable_df["Seleccionar"]]
        st.info(f"Seleccionados: {len(seleccionados)}")

        col1, col2, col3 = st.columns(3)

        # 📊 Excel respaldo
        excel_respaldo, nombre_respaldo = generar_excel_respaldo(resultado_final)
        with col1:
            st.download_button(
                "📊 Registros Observados",
                data=excel_respaldo,
                file_name=nombre_respaldo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # 📄 Formato Due Diligence (plantilla real)
        excel_dd, nombre_dd = generar_formato_due_diligence(resultado_final)
        with col2:
            st.download_button(
                "📄 Formato Due Diligence (Oficial)",
                data=excel_dd,
                file_name=nombre_dd,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # 🚀 Envío API
        with col3:
            if len(seleccionados) > 0:
                referencias = seleccionados["REFERENCIA"].tolist()
                excel_rechazo = generar_excel_rechazo(referencias)

                if st.button("🚀 Enviar Rechazo a la API"):
                    with st.spinner("Enviando a la API..."):
                        response = enviar_rechazo_api(excel_rechazo)

                    if response.status_code == 200:
                        st.success("✅ Rechazo enviado correctamente")
                    else:
                        st.error(f"❌ Error API {response.status_code}\n{response.text}")

    else:
        st.warning("No se encontraron registros mayores a 30K.")
