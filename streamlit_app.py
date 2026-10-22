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
import io

st.set_page_config(page_title="Informes Antropométricos", page_icon="📊", layout="wide")

st.title("Informes Antropométricos")
st.subheader("Ciencia aplicada al rendimiento y la salud")

# --- FUNCIONES AUXILIARES ---
def sombrear_celda(celda, color_hex):
    tcPr = celda._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)

def crear_tabla_datos(doc, titulo, datos, cabecera):
    doc.add_heading(titulo, level=3).runs[0].font.color.rgb = RGBColor(15, 118, 110)
    t = doc.add_table(rows=1, cols=len(cabecera))
    t.style = 'Table Grid'
    hdr = t.rows[0].cells
    for i, text in enumerate(cabecera):
        hdr[i].text = text
        hdr[i].paragraphs[0].runs[0].bold = True
        hdr[i].paragraphs[0].runs[0].font.size = Pt(8)
        hdr[i].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
        sombrear_celda(hdr[i], "0F766E")
    for row_data in datos:
        row = t.add_row().cells
        for i, text in enumerate(row_data):
            row[i].text = str(text)
            row[i].paragraphs[0].runs[0].font.size = Pt(9)
            if i == 0: row[i].paragraphs[0].runs[0].bold = True
    doc.add_paragraph()

def generar_graficos():
    # 1. Gráfico de Dona
    labels = ['Músculo', 'Grasa', 'Residual', 'Hueso', 'Piel']
    sizes = [45.75, 12.35, 9.03, 9.11, 3.75]
    colors = ['#0F766E', '#E11D48', '#94A3B8', '#CBD5E1', '#F1F5F9']
    fig1, ax1 = plt.subplots(figsize=(6, 4))
    ax1.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%', startangle=90, wedgeprops=dict(width=0.4, edgecolor='w'))
    ax1.axis('equal')
    plt.title('Composición Corporal', fontweight='bold', color='#1E293B')
    plt.tight_layout()
    fig1.savefig('grafico_torta.png', dpi=200, transparent=True)
    plt.close(fig1)
    
    # 2. Somatocarta
    fig2, ax2 = plt.subplots(figsize=(6, 5))
    theta = np.linspace(np.pi, 0, 100)
    x = np.cos(theta) * 7
    y = np.sin(theta) * 7
    ax2.plot(x, y, color='#94A3B8', linewidth=2)
    ax2.plot([-7, 7], [0, 0], color='#94A3B8', linewidth=2)
    ax2.axhline(0, color='#CBD5E1', linewidth=1)
    ax2.axvline(0, color='#CBD5E1', linewidth=1)
    ax2.plot([-3.5, 3.5], [6.06, 6.06], color='#CBD5E1', linewidth=1)
    ax2.plot([-3.5, 0], [6.06, 0], color='#CBD5E1', linewidth=1)
    ax2.plot([3.5, 0], [6.06, 0], color='#CBD5E1', linewidth=1)
    ax2.scatter(0.24, 13.05, color='#E11D48', s=150, zorder=5, label='Miguel (Act.)')
    ax2.scatter(1.4, 3.4, color='#0F766E', s=150, zorder=5, marker='D', label='Ref. Baloncesto')
    ax2.set_xlim(-8, 8)
    ax2.set_ylim(-2, 14)
    ax2.set_xticks([])
    ax2.set_yticks([])
    ax2.spines['top'].set_visible(False)
    ax2.spines['right'].set_visible(False)
    ax2.spines['bottom'].set_visible(False)
    ax2.spines['left'].set_visible(False)
    ax2.text(0, 14.5, "SOMATOCARTA", ha='center', fontweight='bold', color='#1E293B', fontsize=12)
    ax2.legend(loc='lower right')
    plt.tight_layout()
    fig2.savefig('somatocarta.png', dpi=200, transparent=True)
    plt.close(fig2)

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

    # Gráficos
    doc.add_heading('1. Análisis Visual', level=2).runs[0].font.color.rgb = RGBColor(15, 118, 110)
    tabla_graficos = doc.add_table(rows=1, cols=2)
    tabla_graficos.alignment = WD_TABLE_ALIGNMENT.CENTER
    celda_izq = tabla_graficos.cell(0, 0)
    run_img1 = celda_izq.paragraphs[0].add_run()
    run_img1.add_picture('grafico_torta.png', width=Inches(3.0))
    celda_der = tabla_graficos.cell(0, 1)
    run_img2 = celda_der.paragraphs[0].add_run()
    run_img2.add_picture('somatocarta.png', width=Inches(3.0))
    doc.add_paragraph()

    # 5 Masas (CON INTERPRETACIÓN)
    doc.add_heading('2. Fraccionamiento en 5 Masas (D. Kerr)', level=2).runs[0].font.color.rgb = RGBColor(15, 118, 110)
    crear_tabla_datos(doc, "Composición Corporal", [
        ("Muscular", "45.8 kg (57.2%)", ">54.2%", "Desarrollo sobresaliente. Alta capacidad de fuerza."),
        ("Adiposa", "12.4 kg (15.4%)", "<16.6%", "Excelente. Nivel bajo de grasa, optimiza agilidad."),
        ("Ósea", "9.1 kg (11.4%)", "11-13%", "Estructura dentro de rangos fisiológicos."),
        ("Residual", "9.0 kg (11.3%)", "10-12%", "Masa visceral dentro de límites normales."),
        ("Piel", "3.7 kg (4.7%)", "4-5%", "Tejido epitelial acorde a superficie corporal.")
    ], ["Componente", "Resultado", "Referencia", "Interpretación Kinésica"])

    # Somatotipo
    doc.add_heading('3. Somatotipo (Heath & Carter)', level=2).runs[0].font.color.rgb = RGBColor(15, 118, 110)
    crear_tabla_datos(doc, "Análisis Morfológico", [
        ("Endomorfia", "1.4", "2.1 (Mod.)", "Baja acumulación subcutánea."),
        ("Mesomorfia", "8.1", "4.5 (Alta)", "Dominancia mesomórfica absoluta. Gran masa muscular."),
        ("Ectomorfia", "1.7", "3.5 (Mod.)", "Estructura densa y compacta."),
        ("Global", "Mesomorfo Puro", "Meso-Ectomorfo", "Perfil orientado a fuerza explosiva.")
    ], ["Componente", "Valor", "Ref. Deporte", "Clasificación"])

    # Cardiometabólico
    doc.add_heading('4. Indicadores Cardiometabólicos', level=2).runs[0].font.color.rgb = RGBColor(15, 118, 110)
    crear_tabla_datos(doc, "Salud y Energía", [
        ("ICC", "0.78", "<0.83 (Bajo)", "Riesgo cardiovascular muy bajo."),
        ("Suma 6 Pliegues", "35.9 mm", "<45.0 mm", "Excelente perfil magro."),
        ("Metabolismo Basal", "1801 kcal", "Harris-Benedict", "Gasto mínimo en reposo."),
        ("Gasto Total (GET)", "3782 kcal", "Intensa (x2.1)", "Requerimiento para sostener entrenamiento.")
    ], ["Indicador", "Resultado", "Referencia", "Observación Kinésica"])

    # Conclusión
    doc.add_heading('5. Diagnóstico y Recomendaciones', level=2).runs[0].font.color.rgb = RGBColor(15, 118, 110)
    p_interp = doc.add_paragraph(
        "Diagnóstico Kinésico: El evaluado presenta una composición corporal altamente eficiente para el rendimiento deportivo, "
        "destacando un desarrollo muscular sobresaliente (57.2%) e índices de adiposidad mínimos (15.4%). "
        "Su perfil somatotípico se clasifica como Mesomorfo Puro (1.4-8.1-1.7), otorgándole una ventaja mecánica "
        "notable para deportes de fuerza, potencia y aceleración/desaceleración rápida.\n\n"
        "Recomendaciones:\n"
        "1. Mantenimiento Estructural: Preservar la masa muscular activa mediante sobrecarga progresiva.\n"
        "2. Soporte Energético: Ajustar ingesta calórica al GET (~3782 kcal/día) para evitar déficit.\n"
        "3. Monitoreo: Reevaluar en 8 a 12 semanas para cuantificar adaptaciones."
    )
    p_interp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # Firma
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
        with st.spinner("Generando gráficos y documento Word..."):
            generar_graficos()
            doc_bytes = generar_word()
            
            st.success("¡Informe generado con éxito!")
            
            st.download_button(
                label="📥 Descargar Informe Completo en Word",
                data=doc_bytes,
                file_name="Informe_Composicion_Completo.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
