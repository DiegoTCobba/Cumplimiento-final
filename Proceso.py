import streamlit as st
import pandas as pd
import requests
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# ===============================
# CONFIGURACIÓN GENERAL
# ===============================
st.set_page_config(
    page_title="Cumplimiento – Rechazos BCP",
    layout="wide"
)

API_URL = "https://q6capnpv09.execute-api.us-east-1.amazonaws.com/OPS/kpayout/v1/payout_process/reject_invoices_batch"

HEADERS = {
    # 🔐 COPIA AQUÍ LOS HEADERS DE POSTMAN (Authorization, etc.)
    # Ejemplo:
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

    # Ajustar formato Excel
    wb = load_workbook(buffer)
    ws = wb.active

    # Forzar Referencia como texto
    for col in ws.iter_cols(min_col=1, max_col=1, min_row=2):
        for cell in col:
            cell.number_format = "@"

    # Ajustar ancho de columnas
    for column_cells in ws.columns:
        max_length = 0
        col_letter = get_column_letter(column_cells[0].column)
        for cell in column_cells:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_length + 3

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


# ===============================
# INTERFAZ
# ===============================
st.title("🚨 Cumplimiento – Rechazo de Clientes (>30K)")
st.write("Carga archivos Excel, selecciona clientes y ejecuta rechazo individual o masivo.")

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

            columnas_interes = df.iloc[:, [1, 2, 3, 8, 12]].copy()
            columnas_interes.columns = [
                "DOCUMENTO",
                "NUMERO_DOCUMENTO",
                "NOMBRE",
                "REFERENCIA",
                "MONTO"
            ]

            columnas_interes["MONTO"] = pd.to_numeric(
                columnas_interes["MONTO"], errors="coerce"
            )

            columnas_interes["REFERENCIA"] = columnas_interes["REFERENCIA"].astype(str)

            filtrado = columnas_interes[columnas_interes["MONTO"] > 30000]
            filtrado["Archivo_Origen"] = file.name

            dataframes.append(filtrado)

        except Exception as e:
            st.error(f"Error procesando {file.name}: {e}")

    if dataframes:
        resultado_final = pd.concat(dataframe_
