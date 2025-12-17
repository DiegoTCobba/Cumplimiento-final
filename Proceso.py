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
    # "Authorization": "Bearer TU_TOKEN"
}

CODIGO_RECHAZO = "R001"
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
st.write("Carga archivos Excel, selecciona clientes y ejecuta rechazo masivo.")

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
            use_container_width=True,
            hide_index=True
        )

        seleccionados = editable_df[editable_df["Seleccionar"]]

        st.info(f"Seleccionados: {len(seleccionados)}")

        if len(seleccionados) > 0:
            referencias = seleccionados["REFERENCIA"].tolist()
            excel_rechazo = generar_excel_rechazo(referencias)

            col1, col2 = st.columns(2)

            with col1:
                st.download_button(
                    "📥 Descargar Excel de Rechazo",
                    data=excel_rechazo,
                    file_name="Rechazo.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            with col2:
                if st.button("🚀 Enviar Rechazo a la API"):
                    with st.spinner("Enviando a la API..."):
                        response = enviar_rechazo_api(excel_rechazo)

                    if response.status_code == 200:
                        st.success("✅ Rechazo enviado correctamente")
                    else:
                        st.error(
                            f"❌ Error API ({response.status_code})\n{response.text}"
                        )

    else:
        st.warning("No se encontraron registros mayores a 30K.")
