import streamlit as st
import pandas as pd
import requests
from datetime import datetime
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
    # "Authorization": "Bearer TU_TOKEN_REAL"
}

# 🔴 EXACTO A POSTMAN
CODIGO_RECHAZO = "001"
DESCRIPCION_RECHAZO = "CUENTA INVALIDA"

# ===============================
# CARGA DE EXCELS
# ===============================
st.title("🚨 Cumplimiento – Clientes Observados (>30K)")

uploaded_files = st.file_uploader(
    "📂 Cargar uno o más archivos Excel",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

dataframes = []

if uploaded_files:
    for file in uploaded_files:
        try:
            df = pd.read_excel(file)

            columnas_interes = df.iloc[:, [1, 2, 3, 8, 12]].copy()
            columnas_interes.columns = [
                "DOCUMENTO",          # DNI / RUC
                "NUMERO_DOCUMENTO",
                "NOMBRE",
                "REFERENCIA",
                "MONTO"
            ]

            columnas_interes["MONTO"] = pd.to_numeric(
                columnas_interes["MONTO"], errors="coerce"
            )

            columnas_interes["REFERENCIA"] = columnas_interes["REFERENCIA"].astype(str)
            columnas_interes["Archivo_Origen"] = file.name

            dataframes.append(columnas_interes)

        except Exception as e:
            st.error(f"❌ Error procesando {file.name}: {e}")

if not dataframes:
    st.stop()

df_total = pd.concat(dataframes, ignore_index=True)

# ===============================
# FILTRO > 30K
# ===============================
resultado_final = df_total[df_total["MONTO"] > 30000].copy()
resultado_final["Seleccionar"] = False

# ===============================
# TABLA STREAMLIT
# ===============================
st.subheader("📋 Clientes detectados")

edited_df = st.data_editor(
    resultado_final,
    use_container_width=True,
    hide_index=True
)

# ===============================
# EXCEL DE EVIDENCIAS (TODOS >30K)
# ===============================
def generar_excel_evidencias(df):
    buffer = BytesIO()
    df.drop(columns=["Seleccionar"]).to_excel(buffer, index=False)
    buffer.seek(0)

    wb = load_workbook(buffer)
    ws = wb.active

    for col in ws.columns:
        max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = max_len + 6

    final_buffer = BytesIO()
    wb.save(final_buffer)
    final_buffer.seek(0)
    return final_buffer

st.download_button(
    "⬇️ Descargar Excel de Evidencias (Todos >30K)",
    generar_excel_evidencias(resultado_final),
    file_name="Evidencias_Clientes_Observados_30K.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# ===============================
# DUE DILIGENCE (PLANTILLA)
# ===============================
def generar_due_diligence(df):
    wb = load_workbook("plantillas/Formato_Due_Diligence_Template.xlsx")
    ws = wb.active

    fila = 13  # inicio real de tabla en la plantilla

    for _, row in df.iterrows():
        ws[f"C{fila}"] = row["DOCUMENTO"]          # Tipo identificación
        ws[f"D{fila}"] = row["NUMERO_DOCUMENTO"]   # Número
        ws[f"E{fila}"] = row["NOMBRE"]              # Razón social
        fila += 1

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    nombre = f"Formato_Due_Diligence_{datetime.now().strftime('%d.%m.%y')}.xlsx"
    return buffer, nombre

buffer_dd, nombre_dd = generar_due_diligence(resultado_final)

st.download_button(
    "⬇️ Descargar Formato Due Diligence",
    buffer_dd,
    file_name=nombre_dd,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# ===============================
# RECHAZO VÍA API (SOLO SELECCIONADOS)
# ===============================
def generar_excel_rechazo(df):
    buffer = BytesIO()

    df_rechazo = pd.DataFrame({
        "Referencia": df["REFERENCIA"].astype(str),
        "Estado": ["Rechazada"] * len(df),
        "Codigo de Rechazo": [CODIGO_RECHAZO] * len(df),
        "Descripcion de Rechazo": [DESCRIPCION_RECHAZO] * len(df)
    })

    df_rechazo.to_excel(
        buffer,
        index=False,
        engine="openpyxl"
    )

    buffer.seek(0)
    return buffer

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

seleccionados = edited_df[edited_df["Seleccionar"]]

# ===============================
# EXPORTAR EXCEL PARA POSTMAN
# ===============================
st.subheader("📄 Excel de Rechazo (Postman)")

if len(seleccionados) > 0:
    excel_rechazo = generar_excel_rechazo(seleccionados)

    st.download_button(
        label="⬇️ Descargar Excel de Rechazo (Postman)",
        data=excel_rechazo,
        file_name="RechazoBCP.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Selecciona clientes para generar el Excel de rechazo.")

#-------------------
if st.button("🚀 Enviar Rechazo a la API"):
    if len(seleccionados) == 0:
        st.warning("Selecciona al menos un cliente.")
    else:
        excel_api = generar_excel_rechazo(seleccionados)
        response = enviar_rechazo_api(excel_api)

        if response.status_code == 200:
            st.success("✅ Rechazo enviado correctamente a la API.")
        else:
            st.error(f"❌ Error API: {response.status_code}")
