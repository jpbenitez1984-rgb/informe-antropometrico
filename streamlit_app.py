import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

st.set_page_config(page_title="Informes Antropométricos", page_icon="📊", layout="wide")

st.title("Informes Antropométricos")
st.subheader("Ciencia aplicada al rendimiento y la salud")

def sombrear_celda(celda, color_hex):
    tcPr = celda._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)

def generar_word():
    doc = Document()
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.4)
        section.bottom_margin = Inches(0.4)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)

    # Encabezado
    header = doc.sections[0].header
    header_p = header.paragraphs[0]
    header_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    header_p.text = "\n[INSERTAR LOGO AQUÍ]\n" 
    
    # Título
    tabla_header = doc.add_table(rows=1, cols=1)
    tabla_header.style = 'Table Grid'
    celda_header = tabla_header.cell(0, 0)
    sombrear_celda(celda_header, "0F766E")
    p_titulo = celda_header.paragraphs[0]
    p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_titulo = p_titulo.add_run("INFORME DE COMPOSICIÓN CORPORAL")
    run_titulo.bold = True
    run_titulo.font.size = Pt(16)
    run_titulo.font.color.rgb = RGBColor(255, 255, 255)
    p_sub = celda_header.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = p_sub.add_run("Salud y Rendimiento Funcional | Klgo. Juan Pablo Benítez Cárdenas")
    run_sub.font.size = Pt(8)
    run_sub.font.color.rgb = RGBColor(255, 255, 255)
    doc.add_paragraph()

    # Paciente
    tabla_paciente = doc.add_table(rows=2, cols=4)
    tabla_paciente.style = 'Table Grid'
    datos_paciente = [
        ["PACIENTE", "EDAD", "FECHA", "DEPORTE"],
        ["Miguel Claro", "26.14 años", "11-10-2025", "Baloncesto"]
    ]
    for i, fila in enumerate(datos_paciente):
        for j, texto in enumerate(fila):
            celda = tabla_paciente.cell(i, j)
            celda.text = texto
            celda.paragraphs[0].runs[0].font.size = Pt(9)
            if i == 0: 
                celda.paragraphs[0].runs[0].bold = True
                sombrear_celda(celda, "F8FAFC")
            else:
                celda.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()

    # Conclusión y Firma
    doc.add_heading('1. Interpretación Clínica', level=2).runs[0].font.color.rgb = RGBColor(15, 118, 110)
    p_interp = doc.add_paragraph(
        "Miguel presenta un perfil de composición corporal de élite para el Baloncesto (Mesomorfia 8.1). "
        "Su % de grasa es excelente (15.4%) y su índice Adiposo/Muscular es óptimo. "
        "El foco debe estar en mantener su masa muscular de 45.7 kg y ajustar la ingesta calórica a su gasto diario de 3782 kcal."
    )
    p_interp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    doc.add_paragraph("\n")
    p_firma_espacio = doc.add_paragraph("[INSERTAR FIRMA AQUÍ]")
    p_firma_espacio.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    p_nombre = doc.add_paragraph()
    p_nombre.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_nombre = p_nombre.add_run("Klgo. Juan Pablo Benítez Cárdenas\n")
    run_nombre.bold = True
    run_nombre.font.size = Pt(10)
    run_cred = p_nombre.add_run("Mg. Fisiología U. de Chile · ISAK Nivel II\n+56 9 4144 0676 | jpbenitez1984@gmail.com")
    run_cred.italic = True
    run_cred.font.size = Pt(9)

    doc.save("Informe_Miguel.docx")
    with open("Informe_Miguel.docx", "rb") as f:
        return f.read()

# --- INTERFAZ DE STREAMLIT ---
archivo = st.file_uploader("Selecciona el archivo Excel del paciente", type=["xlsx", "xls"])

if archivo is not None:
    st.success(f"Archivo cargado: {archivo.name}")
    
    if st.button("Generar Informe Word"):
        with st.spinner("Generando el documento Word..."):
            doc_bytes = generar_word()
            st.success("¡Informe generado con éxito!")
            st.download_button(
                label="📥 Descargar Informe en Word",
                data=doc_bytes,
                file_name="Informe_Antropometrico.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
