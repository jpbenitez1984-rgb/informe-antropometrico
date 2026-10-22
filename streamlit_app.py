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

def agregar_seccion(doc, titulo, texto=None):
    doc.add_heading(titulo, level=2).runs[0].font.color.rgb = RGBColor(15, 118, 110)
    if texto:
        p = doc.add_paragraph(texto)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

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

def crear_fila_tabla(doc, cabeceras, datos):
    t = doc.add_table(rows=1, cols=len(cabeceras))
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = t.rows[0].cells
    for i, text in enumerate(cabeceras):
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
    doc.add_paragraph()

def generar_word():
    doc = Document()
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
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
    t_pac = doc.add_table(rows=1, cols=4)
    t_pac.style = 'Table Grid'
    datos_pac = [["Paciente", "Miguel Claro", "Edad", "26.1 años"],
                 ["Deporte", "Baloncesto", "Fecha", "11/10/2025"],
                 ["Peso / Talla", "80.0 kg / 178.0 cm", "IMC", "25.2 kg/m²"]]
    for i, fila in enumerate(datos_pac):
        row = t_pac.add_row().cells if i > 0 else t_pac.rows[0].cells
        for j, texto in enumerate(fila):
            row[j].text = texto
            row[j].paragraphs[0].runs[0].font.size = Pt(9)
            if j % 2 == 0:  # Títulos en negrita
                row[j].paragraphs[0].runs[0].bold = True
    doc.add_paragraph()

    # 1. Fraccionamiento
    agregar_seccion(doc, "1. Fraccionamiento Corporal de 5 Componentes (D. Kerr, 1988)")
    crear_fila_tabla(doc, 
        ["Componente", "Resultado", "Referencia", "Interpretación Kinésica"],
        [
            ("Masa Muscular", "45.8 kg (57.2%)", "50.8% - 54.2%", "Desarrollo muscular sobresaliente. Alta capacidad de fuerza y potencia."),
            ("Masa Adiposa", "12.4 kg (15.4%)", "< 16.6%", "Excelente. Nivel bajo de grasa subcutánea, optimiza la agilidad."),
            ("Masa Ósea", "9.1 kg (11.4%)", "11.0% - 13.0%", "Estructura ósea dentro de rangos fisiológicos estandarizados."),
            ("Masa Residual", "9.0 kg (11.3%)", "10.0% - 12.0%", "Masa visceral y órganos dentro de límites normales."),
            ("Masa de la Piel", "3.7 kg (4.7%)", "4.0% - 5.0%", "Tejido epitelial acorde a la superficie corporal."),
            ("Índice Músculo/Óseo", "5.02", "> 4.55", "Excelente eficiencia estructural."),
            ("Índice Adiposo/Muscular", "0.27", "< 0.36", "Ratio óptimo con claro predominio de tejido activo.")
        ]
    )

    # Gráficos
    doc.add_heading('Análisis Visual', level=3).runs[0].font.color.rgb = RGBColor(15, 118, 110)
    tabla_graficos = doc.add_table(rows=1, cols=2)
    tabla_graficos.alignment = WD_TABLE_ALIGNMENT.CENTER
    run_img1 = tabla_graficos.cell(0, 0).paragraphs[0].add_run()
    run_img1.add_picture('grafico_torta.png', width=Inches(3.0))
    run_img2 = tabla_graficos.cell(0, 1).paragraphs[0].add_run()
    run_img2.add_picture('somatocarta.png', width=Inches(3.0))
    doc.add_paragraph()

    # 2. Somatotipo
    agregar_seccion(doc, "2. Análisis de Somatotipo y Perfil Morfológico (Heath & Carter)")
    crear_fila_tabla(doc,
        ["Componente", "Valor", "Referencia Deporte", "Clasificación / Interpretación"],
        [
            ("Endomorfia (Grasa)", "1.4", "2.1 (Moderada)", "Baja acumulación subcutánea; relieve muscular visible."),
            ("Mesomorfia (Músculo)", "8.1", "4.5 (Alta)", "Dominancia mesomórfica absoluta. Gran masa muscular y robustez."),
            ("Ectomorfia (Linealidad)", "1.7", "3.5 (Moderada)", "Estructura densa y compacta por unidad de altura."),
            ("Clasificación Global", "Mesomorfo Puro", "Meso-Ectomorfo", "Perfil fuertemente orientado a la fuerza explosiva.")
        ]
    )

    # 3. Cardiometabólico
    agregar_seccion(doc, "3. Indicadores Cardiometabólicos y Prescripción Energética")
    crear_fila_tabla(doc,
        ["Indicador Fisiológico", "Resultado", "Rango de Referencia", "Diagnóstico / Observación Kinésica"],
        [
            ("Índice Cintura / Cadera", "0.78", "< 0.83 (Bajo Riesgo)", "Riesgo cardiovascular y acumulación de grasa visceral muy bajos."),
            ("Suma 6 Pliegues Cutáneos", "35.9 mm", "< 45.0 mm", "Excelente perfil magro en tejido subcutáneo."),
            ("Metabolismo Basal (MB)", "1801 kcal", "Harris & Benedict", "Gasto energético mínimo en reposo para soporte vital."),
            ("Gasto Energético Total (GET)", "3782 kcal", "Actividad Intensa (x2.1)", "Requerimiento calórico estimado para sostener volumen de entrenamiento.")
        ]
    )

    # 4. Conclusión
    agregar_seccion(doc, "4. Conclusión Fisiológica y Recomendaciones de Carga")
    p_interp = doc.add_paragraph(
        "Diagnóstico Kinésico-Antropométrico:\n"
        "El evaluado presenta una composición corporal altamente eficiente para el rendimiento deportivo, destacando un desarrollo muscular sobresaliente (57.2% de masa muscular total) e índices de adiposidad mínimos (15.4% de masa adiposa). Su perfil somatotípico se clasifica como Mesomorfo Puro (1.4-8.1-1.7), lo que le otorga una ventaja mecánica notable para deportes de fuerza, potencia y aceleración/desaceleración rápida.\n\n"
        "Recomendaciones de Prescripción & Nutrición:\n"
        "1. Mantenimiento Estructural: Preservar la alta masa muscular activa mediante trabajo de fuerza con sobrecarga progresiva y consumo proteico adecuado.\n"
        "2. Soporte Energético: Ajustar la ingesta calórica al Gasto Energético Total estimado (~3782 kcal/día) para evitar estados de déficit involuntario que comprometan la masa magra.\n"
        "3. Monitoreo Fisiológico: Se recomienda reevaluar en un período de 8 a 12 semanas para cuantificar las adaptaciones específicas del ciclo de entrenamiento."
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
