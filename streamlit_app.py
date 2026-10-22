
import streamlit as st

st.set_page_config(
    page_title="Informes antropométricos",
    page_icon="📊",
    layout="wide"
)

st.title("Informes antropométricos")
st.subheader("Ciencia aplicada al rendimiento y la salud")

st.write(
    "La aplicación está funcionando. "
    "Próximamente podrás cargar archivos Excel y generar informes."
)

archivo = st.file_uploader(
    "Selecciona un archivo Excel",
    type=["xlsx", "xlsm", "xls"]
)

if archivo is not None:
    st.success(f"Archivo seleccionado: {archivo.name}")
