import streamlit as st
import pandas as pd
import requests
from io import BytesIO
from datetime import datetime
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

    for column_cells in ws.columns:
        max_length = 0
        col_letter = get_column_letter(column_cells[0].column)
        for cell in column_cells:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_length + 4

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


def generar_formato_due_diligence(df):
    """
    Llena la plantilla EXACTAMENTE en la tabla azul
    Corrige el tipo de identificación (RUC)
    """
    wb = load_workbook("plantillas/Formato_Due_Diligence_Template.xlsx")
    ws = wb.active

    # Encabezado fijo
    ws["C9"] = "Operaciones"
    ws["C11"] = datetime.now().strftime("%d/%m/%Y")

    # La tabla azul empieza en la fila 13
    fila = 13

    for _, row in df.iterrows():
        ws[f"B{fila}"] = "RUC"                          # ✅ FORZADO
        ws[f"C{fila}"] = str(row["NUMERO_DOCUMENTO"])   # Número
        ws[f"D{fila}"] = row["NOMBRE"]                  # Razón Social
        fila += 1

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    fecha_archivo = datetime.now().strftime("%d.%m.%y")
    nombre_archivo = f"Formato_Due_Diligence_{fecha_archivo}.xlsx"

    return buffer, nombre_archivo


# ===============================
# INTERFAZ
# ===============================
st.title("🚨 Cumplimiento – Rechazo de Clientes (>30K)")
st.write("Carga archivos Excel, revisa clientes observados, genera evidencias, Due Diligence y rechazo.")

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
        resultado_final = pd.concat(dataframes, ignore_index=True)

        resultado_final.insert(0, "Seleccionar", False)

        st.subheader("📋 Clientes Observados (>30K)")
        edited_df = st.data_editor(
            resultado_final,
            use_container_width=True,
            num_rows="dynamic"
        )

        seleccionados = edited_df[edited_df["Seleccionar"]]

        # ===============================
        # EXCEL DE EVIDENCIAS (TODOS >30K)
        # ===============================
        buffer_evidencias = BytesIO()
        resultado_final.drop(columns=["Seleccionar"]).to_excel(
            buffer_evidencias,
            index=False,
            engine="openpyxl"
        )
        buffer_evidencias.seek(0)

        wb = load_workbook(buffer_evidencias)
        ws = wb.active

        for column_cells in ws.columns:
            max_length = 0
            col_letter = get_column_letter(column_cells[0].column)
            for cell in column_cells:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = max_length + 5

        buffer_evidencias_final = BytesIO()
        wb.save(buffer_evidencias_final)
        buffer_evidencias_final.seek(0)

        st.download_button(
            "⬇️ Descargar Excel de Evidencias (Todos >30K)",
            data=buffer_evidencias_final,
            file_name="Evidencias_Clientes_Observados_30K.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # ===============================
        # DUE DILIGENCE (TODOS >30K)
        # ===============================
        excel_dd, nombre_dd = generar_formato_due_diligence(resultado_final)

        st.download_button(
            "📄 Descargar Formato Due Diligence (Todos >30K)",
            data=excel_dd,
            file_name=nombre_dd,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # ===============================
        # RECHAZO VIA POSTMAN (SELECCIONADOS)
        # ===============================
        if not seleccionados.empty:
            st.subheader("🚫 Rechazo vía Postman / API")

            if st.button("Ejecutar Rechazo de Clientes Seleccionados"):
                referencias = seleccionados["REFERENCIA"].tolist()
                excel_api = generar_excel_rechazo(referencias)
                response = enviar_rechazo_api(excel_api)

                if response.status_code in (200, 201):
                    st.success("✅ Rechazo ejecutado correctamente.")
                else:
                    st.error("❌ Error en rechazo vía API.")
                    st.write("Status:", response.status_code)
                    st.text(response.text)

