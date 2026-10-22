import streamlit as st
import matplotlib.pyplot as plt
import numpy as np
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io

st.set_page_config(page_title="Dossier Kinesiológico", page_icon="📊", layout="wide")

st.title("Generador de Dossier Kinesiológico")
st.subheader("Evaluación de Composición Corporal ISAK II")

def set_cell_border(cell, **kwargs):
    """Aplica bordes específicos a celdas (para quitar los negros feos de Word)."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    for edge in ('start', 'top', 'end', 'bottom', 'insideH', 'insideV'):
        if edge in kwargs:
            edge_data = kwargs.get(edge)
            tag = 'w:{}'.format(edge)
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)
            for key in ["sz", "val", "color", "space"]:
                if key in edge_data:
                    element.set(qn('w:{}'.format(key)), str(edge_data[key]))

def sombrear_celda(celda, color_hex):
    tcPr = celda._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)

def generar_graficos():
    labels = ['Músculo', 'Grasa', 'Residual', 'Hueso', 'Piel']
    sizes = [45.75, 12.35, 9.03, 9.11, 3.75]
    colors = ['#0F766E', '#E11D48', '#94A3B8', '#CBD5E1', '#F1F5F9']
    fig1, ax1 = plt.subplots(figsize=(6, 4))
    ax1.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%', startangle=90, wedgeprops=dict(width=0.4, edgecolor='w'))
    ax1.axis('equal')
    plt.title('Composición Corporal (Kerr)', fontweight='bold', color='#1E293B', fontsize=10)
    plt.tight_layout()
    fig1.savefig('grafico_torta.png', dpi=300, transparent=True)
    plt.close(fig1)
    
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
    plt.title('Somatocarta', fontweight='bold', color='#1E293B', fontsize=10)
    ax2.legend(loc='lower right', fontsize=8)
    plt.tight_layout()
    fig2.savefig('somatocarta.png', dpi=300, transparent=True)
    plt.close(fig2)

def crear_tabla_estilizada(doc, headers, data):
    t = doc.add_table(rows=1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Encabezados
    hdr = t.rows[0].cells
    for i, text in enumerate(headers):
        hdr[i].text = text
        hdr[i].paragraphs[0].runs[0].bold = True
        hdr[i].paragraphs[0].runs[0].font.size = Pt(8)
        hdr[i].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
        sombrear_celda(hdr[i], "0F766E")
        # Bordes solo inferior
        set_cell_border(hdr[i], bottom={"val": "single", "sz": 12, "color": "0F766E"}, top={"val": "nil"}, left={"val": "nil"}, right={"val": "nil"})
        
    # Datos
    for row_data in data:
        row = t.add_row().cells
        for i, text in enumerate(row_data):
            row[i].text = text
            row[i].paragraphs[0].runs[0].font.size = Pt(8)
            if i == 0: row[i].paragraphs[0].runs[0].bold = True
            # Solo borde inferior gris claro
            set_cell_border(row[i], bottom={"val": "single", "sz": 4, "color": "E2E8F0"}, top={"val": "nil"}, left={"val": "nil"}, right={"val": "nil"})
    doc.add_paragraph()

def generar_word():
    doc = Document()
    # Márgenes idénticos al PDF
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)

    # --- ENCABEZADO (Dossier) ---
    header = doc.sections[0].header
    header_p = header.paragraphs[0]
    header_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    header_p.text = "\n[INSERTAR LOGO AQUÍ]\n" 

    # TÍTULO Y DATOS
    tabla_top = doc.add_table(rows=1, cols=2)
    tabla_top.autofit = False
    cell_izq = tabla_top.cell(0, 0)
    cell_izq.width = Inches(4.0)
    p_tit = cell_izq.paragraphs[0]
    run_tit = p_tit.add_run("DOSSIER DE EVALUACIÓN\nKINESIOLÓGICA")
    run_tit.bold = True
    run_tit.font.size = Pt(14)
    run_tit.font.color.rgb = RGBColor(15, 118, 110)
    p_sub = cell_izq.add_paragraph("Composición Corporal & Fisiología Aplicada")
    p_sub.runs[0].font.size = Pt(8)
    p_sub.runs[0].font.color.rgb = RGBColor(100, 116, 139)

    cell_der = tabla_top.cell(0, 1)
    cell_der.width = Inches(3.0)
    p_der = cell_der.paragraphs[0]
    p_der.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_kine = p_der.add_run("Klgo. Juan Pablo Benítez C.\n")
    run_kine.bold = True
    run_kine.font.size = Pt(10)
    run_kine_2 = p_der.add_run("Mg. Fisiología U. de Chile · ISAK II\n+56 9 4144 0676 | jpbenitez1984@gmail.com")
    run_kine_2.font.size = Pt(8)
    run_kine_2.italic = True

    # Línea verde separadora
    p_line = doc.add_paragraph()
    pPr = p_line._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '12')
    bottom.set(qn('w:color'), '0F766E')
    pBdr.append(bottom)
    pPr.append(pBdr)
    doc.add_paragraph()

    # DATOS PACIENTE
    t_pac = doc.add_table(rows=2, cols=3)
    datos_pac = [
        ["Paciente: Miguel Claro", "Edad: 26.1 años", "Deporte: Baloncesto"],
        ["Fecha: 11/10/2025", "Peso / Talla: 80.0 kg / 178.0 cm", "IMC: 25.2 kg/m²"]
    ]
    for i, fila in enumerate(datos_pac):
        for j, texto in enumerate(fila):
            celda = t_pac.cell(i, j)
            celda.text = texto
            celda.paragraphs[0].runs[0].font.size = Pt(8)
            set_cell_border(celda, top={"val": "single", "sz": 4, "color": "E2E8F0"}, bottom={"val": "single", "sz": 4, "color": "E2E8F0"}, left={"val": "single", "sz": 4, "color": "E2E8F0"}, right={"val": "single", "sz": 4, "color": "E2E8F0"})
    doc.add_paragraph()

    # 1. TABLA FRACCIONAMIENTO
    doc.add_heading('1. Fraccionamiento Corporal de 5 Componentes (D. Kerr, 1988)', level=2).runs[0].font.color.rgb = RGBColor(15, 118, 110)
    crear_tabla_estilizada(doc, 
        ["Componente", "Resultado", "Referencia", "Interpretación Kinésica"],
        [
            ("Masa Muscular", "45.8 kg (57.2%)", "50.8% - 54.2%", "Desarrollo muscular sobresaliente. Alta capacidad de fuerza y potencia."),
            ("Masa Adiposa", "12.4 kg (15.4%)", "< 16.6%", "Excelente. Nivel bajo de grasa subcutánea, optimiza la agilidad."),
            ("Masa Ósea", "9.1 kg (11.4%)", "11.0% - 13.0%", "Estructura ósea dentro de rangos fisiológicos."),
            ("Masa Residual", "9.0 kg (11.3%)", "10.0% - 12.0%", "Masa visceral y órganos dentro de límites."),
            ("Masa de la Piel", "3.7 kg (4.7%)", "4.0% - 5.0%", "Tejido epitelial acorde a la superficie."),
            ("Índice Músculo/Óseo", "5.02", "> 4.55", "Excelente eficiencia estructural."),
            ("Índice Adiposo/Muscular", "0.27", "< 0.36", "Ratio óptimo con predominio activo.")
        ]
    )
    
    p_torta = doc.add_paragraph()
    p_torta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_torta.add_run().add_picture('grafico_torta.png', width=Inches(3.5))

    # 2. TABLA SOMATOTIPO
    doc.add_page_break()
    doc.add_heading('2. Análisis de Somatotipo y Perfil Morfológico', level=2).runs[0].font.color.rgb = RGBColor(15, 118, 110)
    crear_tabla_estilizada(doc,
        ["Componente", "Valor", "Ref. Deporte", "Clasificación"],
        [
            ("Endomorfia", "1.4", "2.1 (Mod.)", "Baja acumulación subcutánea; relieve muscular."),
            ("Mesomorfia", "8.1", "4.5 (Alta)", "Dominancia mesomórfica absoluta."),
            ("Ectomorfia", "1.7", "3.5 (Mod.)", "Estructura densa y compacta."),
            ("Global", "Mesomorfo Puro", "Meso-Ectomorfo", "Orientado a fuerza explosiva.")
        ]
    )
    
    p_soma = doc.add_paragraph()
    p_soma.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_soma.add_run().add_picture('somatocarta.png', width=Inches(4.0))

    # 3. TABLA CARDIOMETABÓLICO
    doc.add_heading('3. Indicadores Cardiometabólicos', level=2).runs[0].font.color.rgb = RGBColor(15, 118, 110)
    crear_tabla_estilizada(doc,
        ["Indicador Fisiológico", "Resultado", "Referencia", "Observación Kinésica"],
        [
            ("Índice Cintura / Cadera", "0.78", "< 0.83 (Bajo Riesgo)", "Riesgo cardiovascular muy bajo."),
            ("Suma 6 Pliegues Cutáneos", "35.9 mm", "< 45.0 mm", "Excelente perfil magro."),
            ("Metabolismo Basal", "1801 kcal", "Harris & Benedict", "Gasto mínimo en reposo."),
            ("Gasto Energético Total", "3782 kcal", "Actividad Intensa", "Requerimiento para entrenamiento.")
        ]
    )

    # 4. CONCLUSIÓN CON RECUADRO
    doc.add_heading('4. Conclusión Fisiológica y Recomendaciones', level=2).runs[0].font.color.rgb = RGBColor(15, 118, 110)
    
    tabla_concl = doc.add_table(rows=1, cols=1)
    cell_concl = tabla_concl.cell(0, 0)
    # Borde verde grueso
    set_cell_border(cell_concl, 
        top={"val": "single", "sz": 12, "color": "0F766E"}, bottom={"val": "single", "sz": 12, "color": "0F766E"},
        left={"val": "single", "sz": 12, "color": "0F766E"}, right={"val": "single", "sz": 12, "color": "0F766E"}
    )
    
    p_diag = cell_concl.paragraphs[0]
    run_diag = p_diag.add_run("Diagnóstico Kinésico-Antropométrico:\n")
    run_diag.bold = True
    run_diag.font.size = Pt(9)
    run_diag.font.color.rgb = RGBColor(15, 118, 110)
    
    p_texto = cell_concl.add_paragraph("El evaluado presenta una composición corporal altamente eficiente para el rendimiento deportivo, destacando un desarrollo muscular sobresaliente (57.2%) e índices de adiposidad mínimos (15.4%).")
    p_texto.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_texto.runs[0].font.size = Pt(9)

    p_recom = cell_concl.add_paragraph()
    run_recom = p_recom.add_run("\nRecomendaciones de Prescripción & Nutrición:\n")
    run_recom.bold = True
    run_recom.font.size = Pt(9)
    run_recom.font.color.rgb = RGBColor(15, 118, 110)
    
    p_items = cell_concl.add_paragraph("1. Mantenimiento Estructural: Preservar masa muscular.\n2. Soporte Energético: Ajustar a GET (~3782 kcal/día).\n3. Monitoreo: Reevaluar en 8-12 semanas.")
    p_items.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_items.runs[0].font.size = Pt(9)

    # Firma
    doc.add_paragraph("\n")
    p_firma = doc.add_paragraph("[INSERTAR FIRMA AQUÍ]")
    p_firma.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    p_nombre = doc.add_paragraph()
    p_nombre.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_nombre = p_nombre.add_run("Klgo. Juan Pablo Benítez Cárdenas\n")
    run_nombre.bold = True
    run_nombre.font.size = Pt(10)
    run_cred = p_nombre.add_run("Mg. Fisiología U. de Chile · ISAK Nivel II\n+56 9 4144 0676 | jpbenitez1984@gmail.com")
    run_cred.italic = True
    run_cred.font.size = Pt(9)

    doc.save("Dossier_Kinesiologico.docx")
    with open("Dossier_Kinesiologico.docx", "rb") as f:
        return f.read()

# --- INTERFAZ ---
archivo = st.file_uploader("Sube el Excel del paciente", type=["xlsx", "xls"])

if archivo is not None:
    if st.button("Generar Dossier"):
        with st.spinner("Generando..."):
            generar_graficos()
            doc_bytes = generar_word()
            st.success("¡Dossier listo!")
            st.download_button(
                label="📥 Descargar Dossier en Word",
                data=doc_bytes,
                file_name="Dossier_Kinesiologico.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
