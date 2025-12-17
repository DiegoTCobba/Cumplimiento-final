import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# =========================
# CONFIG
# =========================
API_URL = "http://TU_API_RECHAZO_URL"

st.set_page_config(page_title="Unidad de Cumplimiento", layout="wide")

# =========================
# DATA DE EJEMPLO (REEMPLAZA POR TU PIPELINE REAL)
# =========================
data = [
    {
        "DOCUMENTO": "RUC",
        "NUMERO_DOCUMENTO": "20612550264",
        "NOMBRE": "Corporacion ALV EIRL NULL",
        "REFERENCIA": "253506686395",
        "MONTO": 39500.90,
        "Archivo_Origen": "VARIOS PAYOUT IBK"
    },
    {
        "DOCUMENTO": "RUC",
        "NUMERO_DOCUMENTO": "20602935559",
        "NOMBRE": "Motoservice SAC NULL",
        "REFERENCIA": "253506686409",
        "MONTO": 37931.06,
        "Archivo_Origen": "VARIOS PAYOUT BBVA"
    }
]

df = pd.DataFrame(data)

# =========================
# FILTRO > 30K
# =========================
resultado_final = df[df["MONTO"] > 30000].copy()
resultado_final["Seleccionar"] = False

# =========================
# STREAMLIT TABLE
# =========================
st.title("📋 Registros detectados (>30K)")

edited_df = st.data_editor(
    resultado_final,
    use_container_width=True,
    hide_index=True
)

# =========================
# EXCEL DE EVIDENCIAS (TODOS >30K)
# =========================
def generar_excel_evidencias(df):
    buffer = BytesIO()
    df.drop(columns=["Seleccionar"]).to_excel(buffer, index=False)
    buffer.seek(0)

    wb = load_workbook(buffer)
    ws = wb.active

    for col in ws.columns:
        max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = max_len + 5

    final_buffer = BytesIO()
    wb.save(final_buffer)
    final_buffer.seek(0)
    return final_buffer

st.download_button(
    "⬇️ Descargar Excel de Evidencias",
    generar_excel_evidencias(resultado_final),
    file_name="Evidencias_Clientes_Observados_30K.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# =========================
# DUE DILIGENCE (PLANTILLA)
# =========================
def generar_due_diligence(df):
    wb = load_workbook("plantillas/Formato_Due_Diligence_Template.xlsx")
    ws = wb.active

    ws["C9"] = "Operaciones"
    ws["C11"] = datetime.now().strftime("%d/%m/%Y")

    fila = 13  # inicio real de tabla azul

    for _, row in df.iterrows():
        ws[f"A{fila}"] = row["DOCUMENTO"]          # RUC / DNI
        ws[f"B{fila}"] = row["NUMERO_DOCUMENTO"]
        ws[f"C{fila}"] = row["NOMBRE"]
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

# =========================
# RECHAZO API (SOLO SELECCIONADOS)
# =========================
def generar_excel_api(df):
    columnas_api = ["DOCUMENTO", "NUMERO_DOCUMENTO", "NOMBRE", "REFERENCIA", "MONTO"]
    buffer = BytesIO()
    df[columnas_api].to_excel(buffer, index=False)
    buffer.seek(0)
    return buffer

def enviar_rechazo_api(excel_buffer):
    files = {
        "file": (
            "rechazo.xlsx",
            excel_buffer,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    }
    response = requests.post(API_URL, files=files, timeout=60)
    return response

seleccionados = edited_df[edited_df["Seleccionar"]]

if st.button("🚀 Enviar Rechazo a la API"):
    if len(seleccionados) == 0:
        st.warning("Selecciona al menos un cliente.")
    else:
        excel_api = generar_excel_api(seleccionados)
        response = enviar_rechazo_api(excel_api)

        if response.status_code == 200:
            st.success("✅ Rechazo enviado correctamente.")
        else:
            st.error(f"❌ Error API: {response.status_code}")
