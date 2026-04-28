import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="CRM Sencillo",
    page_icon="",
    layout="wide"
)

ESTADOS = [
    "Nueva",
    "Contacto",
    "Negociacion",
    "Ganada",
    "Perdida"
]

COLUMNAS = [
    "Estado",
    "Nombre de Oportunidad",
    "Fecha",
    "Cliente",
    "Representante",
    "Costo de Venta",
    "Costo Final",
    "Fecha de cierre",
    "Comentarios"
]

def crear_tabla_vacia():
    return pd.DataFrame(columns=COLUMNAS)

def normalizar_tabla(df):
    for columna in COLUMNAS:
        if columna not in df.columns:
            df[columna] = None

    df = df[COLUMNAS]

    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
    df["Fecha de cierre"] = pd.to_datetime(df["Fecha de cierre"], errors="coerce")

    df["Costo de Venta"] = pd.to_numeric(df["Costo de Venta"], errors="coerce")
    df["Costo Final"] = pd.to_numeric(df["Costo Final"], errors="coerce")

    df["Estado"] = df["Estado"].fillna("Nueva")

    return df

def convertir_excel(df):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="CRM")

        worksheet = writer.sheets["CRM"]

        for col in worksheet.columns:
            max_length = 0
            column_letter = col[0].column_letter

            for cell in col:
                try:
                    value_length = len(str(cell.value))
                    if value_length > max_length:
                        max_length = value_length
                except:
                    pass

            worksheet.column_dimensions[column_letter].width = max_length + 3

    output.seek(0)
    return output

if "crm_data" not in st.session_state:
    st.session_state.crm_data = crear_tabla_vacia()

st.title("People CRM")
st.caption("CRM sencillo estilo monday.com para capturar oportunidades, importar archivos y exportar datos.")

st.divider()

archivo = st.file_uploader(
    "Importar archivo CSV o Excel",
    type=["csv", "xlsx"]
)

if archivo is not None:
    if st.button("Importar datos"):
        try:
            if archivo.name.endswith(".csv"):
                df_importado = pd.read_csv(archivo)
            else:
                df_importado = pd.read_excel(archivo)

            st.session_state.crm_data = normalizar_tabla(df_importado)
            st.success("Datos importados correctamente.")

        except Exception as e:
            st.error(f"No se pudo importar el archivo: {e}")

st.divider()

st.subheader("Oportunidades")

tabla_editada = st.data_editor(
    st.session_state.crm_data,
    use_container_width=True,
    hide_index=True,
    num_rows="dynamic",
    column_config={
        "Estado": st.column_config.SelectboxColumn(
            "Estado",
            options=ESTADOS,
            required=True,
            width="medium"
        ),
        "Nombre de Oportunidad": st.column_config.TextColumn(
            "Nombre de Oportunidad",
            width="large"
        ),
        "Fecha": st.column_config.DateColumn(
            "Fecha",
            format="DD/MM/YYYY"
        ),
        "Cliente": st.column_config.TextColumn(
            "Cliente",
            width="medium"
        ),
        "Representante": st.column_config.TextColumn(
            "Representante",
            width="medium"
        ),
        "Costo de Venta": st.column_config.NumberColumn(
            "Costo de Venta",
            format="$ %.2f"
        ),
        "Costo Final": st.column_config.NumberColumn(
            "Costo Final",
            format="$ %.2f"
        ),
        "Fecha de cierre": st.column_config.DateColumn(
            "Fecha de cierre",
            format="DD/MM/YYYY"
        ),
        "Comentarios": st.column_config.TextColumn(
            "Comentarios",
            width="large"
        )
    },
    key="editor_crm"
)

tabla_editada["Estado"] = tabla_editada["Estado"].fillna("Nueva")

st.session_state.crm_data = tabla_editada

st.divider()

col1, col2 = st.columns(2)

with col1:
    csv = tabla_editada.to_csv(index=False).encode("utf-8-sig")

    st.download_button(
        label="Descargar CSV",
        data=csv,
        file_name="crm_oportunidades.csv",
        mime="text/csv",
        use_container_width=True
    )

with col2:
    excel = convertir_excel(tabla_editada)

    st.download_button(
        label="Descargar Excel",
        data=excel,
        file_name="crm_oportunidades.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

st.caption("Nota: los datos se mantienen mientras la app está abierta. Para continuar después, descarga el archivo y vuelve a importarlo cuando lo necesites.")
