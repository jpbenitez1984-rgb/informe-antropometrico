import streamlit as st
import matplotlib.pyplot as plt
import numpy as np
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io

st.set_page_config(page_title="Informes Antropométricos", page_icon="📊", layout="wide")

st.title("Informes Antropométricos")
st.subheader("Ciencia aplicada al rendimiento y la salud")

# --- FUNCIONES AUXILIARES ---
def set_cell_border(cell, **kwargs):
    """Dibuja bordes específicos y finos en celdas de Word."""
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

def agregar_linea_verde(p):
    """Agrega un borde inferior verde a un párrafo para simular una línea fina."""
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '12') # Grosor de la línea
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '0F766E')
    pBdr.append(bottom)
    pPr.append(pBdr)

def generar_graficos():
    labels = ['Músculo', 'Grasa', 'Residual', 'Hueso', 'Piel']
    sizes = [45.75, 12.35, 9.03, 9.11, 3.75]
    colors = ['#0F766E', '#E11D48', '#94A3B8', '#CBD5E1', '#F1F5F9']
    fig1, ax1 = plt.subplots(figsize=(6, 4))
    ax1.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%', startangle=90, wedgeprops=dict(width=0.4, edgecolor='w'))
    ax1.axis('equal')
    plt.title('Composición Corporal', fontweight='bold', color='#1E293B', fontsize=10)
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

def generar_word():
    doc = Document()
    # Márgenes más ajustados para que se vea como el PDF
    for section in doc.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)

    # --- ENCABEZADO ---
    header = doc.sections[0].header
    header_p = header.paragraphs[0]
    header_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    header_p.text = "\n[INSERTAR LOGO AQUÍ]\n" 
    
    # 1. TÍTULO Y CONTACTO (Dos columnas)
    tabla_top = doc.add_table(rows=1, cols=2)
    tabla_top.autofit = False
    
    # Celda Izquierda: Título
    cell_izq = tabla_top.cell(0, 0)
    cell_izq.width = Inches(3.5)
    p_tit = cell_izq.paragraphs[0]
    run_tit = p_tit.add_run("DOSSIER DE EVALUACIÓN\nKINESIOLÓGICA")
    run_tit.bold = True
    run_tit.font.size = Pt(14)
    run_tit.font.color.rgb = RGBColor(15, 118, 110)
    p_sub = cell_izq.add_paragraph("Composición Corporal & Fisiología Aplicada")
    p_sub.runs[0].font.size = Pt(8)
    p_sub.runs[0].font.color.rgb = RGBColor(100, 116, 139)

    # Celda Derecha: Datos del Kinesiólogo
    cell_der = tabla_top.cell(0, 1)
    cell_der.width = Inches(3.5)
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
    agregar_linea_verde(p_line)
    doc.add_paragraph()

    # 2. DATOS DEL PACIENTE
    t_pac = doc.add_table(rows=3, cols=3)
    t_pac.style = 'Table Grid'
    t_pac.alignment = WD_TABLE_ALIGNMENT.CENTER
    datos_pac = [
        ["Paciente: Miguel Claro", "Edad: 26.1 años", "Deporte: Baloncesto"],
        ["Fecha: 11/10/2025", "Peso / Talla: 80.0 kg / 178.0 cm", "IMC: 25.2 kg/m²"]
    ]
    for i, fila in enumerate(datos_pac):
        for j, texto in enumerate(fila):
            celda = t_pac.cell(i, j)
            celda.text = texto
            celda.paragraphs[0].runs[0].font.size = Pt(8)
            # Bordes muy sutiles
            set_cell_border(celda, top={"val": "single", "sz": 4, "color": "E2E8F0"}, 
                            bottom={"val": "single", "sz": 4, "color": "E2E8F0"},
                            left={"val": "single", "sz": 4, "color": "E2E8F0"},
                            right={"val": "single", "sz": 4, "color": "E2E8F0"})
    doc.add_paragraph()

    # 3. TABLA 1: FRACCIONAMIENTO
    doc.add_heading('1. Fraccionamiento Corporal de 5 Componentes (D. Kerr, 1988)', level=2).runs[0].font.color.rgb = RGBColor(15, 118, 110)
    
    tabla_1 = doc.add_table(rows=1, cols=4)
    tabla_1.style = 'Table Grid'
    hdr_1 = tabla_1.rows[0].cells
    cabeceras_1 = ["Componente", "Resultado", "Referencia", "Interpretación Kinésica"]
    for i, text in enumerate(cabeceras_1):
        hdr_1[i].text = text
        hdr_1[i].paragraphs[0].runs[0].bold = True
        hdr_1[i].paragraphs[0].runs[0].font.size = Pt(8)
        hdr_1[i].paragraphs[0].runs[0].font.color.rgb = RGBColor(15, 118, 110)
        # Borde inferior grueso verde, resto fino
        set_cell_border(hdr_1[i], bottom={"val": "single", "sz": 12, "color": "0F766E"},
                        top={"val": "nil"}, left={"val": "nil"}, right={"val": "nil"})

    datos_1 = [
        ("Masa Muscular", "45.8 kg (57.2%)", "50.8% - 54.2%", "Desarrollo muscular sobresaliente."),
        ("Masa Adiposa", "12.4 kg (15.4%)", "< 16.6%", "Excelente. Nivel bajo de grasa."),
        ("Masa Ósea", "9.1 kg (11.4%)", "11.0% - 13.0%", "Estructura dentro de rangos."),
        ("Masa Residual", "9.0 kg (11.3%)", "10.0% - 12.0%", "Masa visceral dentro de límites."),
        ("Masa de la Piel", "3.7 kg (4.7%)", "4.0% - 5.0%", "Tejido epitelial acorde."),
        ("Índice Músculo/Óseo", "5.02", "> 4.55", "Excelente eficiencia estructural."),
        ("Índice Adiposo/Muscular", "0.27", "< 0.36", "Ratio óptimo, predominio activo.")
    ]
    for row_data in datos_1:
        row = tabla_1.add_row().cells
        for i, text in enumerate(row_data):
            row[i].text = text
            row[i].paragraphs[0].runs[0].font.size = Pt(8)
            if i == 0: row[i].paragraphs[0].runs[0].bold = True
            set_cell_border(row[i], top={"val": "nil"}, bottom={"val": "single", "sz": 4, "color": "E2E8F0"}, left={"val": "nil"}, right={"val": "nil"})
    
    doc.add_paragraph()

    # Gráfico de Torta
    p_torta = doc.add_paragraph()
    p_torta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_torta.add_run().add_picture('grafico_torta.png', width=Inches(3.5))

    # 4. TABLA 2: SOMATOTIPO
    doc.add_page_break() # Salto de página para que quepa todo bien
    
    doc.add_heading('2. Análisis de Somatotipo y Perfil Morfológico', level=2).runs[0].font.color.rgb = RGBColor(15, 118, 110)
    tabla_2 = doc.add_table(rows=1, cols=4)
    tabla_2.style = 'Table Grid'
    hdr_2 = tabla_2.rows[0].cells
    cabeceras_2 = ["Componente", "Valor", "Ref. Deporte", "Clasificación"]
    for i, text in enumerate(cabeceras_2):
        hdr_2[i].text = text
        hdr_2[i].paragraphs[0].runs[0].bold = True
        hdr_2[i].paragraphs[0].runs[0].font.size = Pt(8)
        hdr_2[i].paragraphs[0].runs[0].font.color.rgb = RGBColor(15, 118, 110)
        set_cell_border(hdr_2[i], bottom={"val": "single", "sz": 12, "color": "0F766E"}, top={"val": "nil"}, left={"val": "nil"}, right={"val": "nil"})

    datos_2 = [
        ("Endomorfia", "1.4", "2.1 (Mod.)", "Baja acumulación subcutánea."),
        ("Mesomorfia", "8.1", "4.5 (Alta)", "Dominancia mesomórfica."),
        ("Ectomorfia", "1.7", "3.5 (Mod.)", "Estructura densa."),
        ("Global", "Mesomorfo Puro", "Meso-Ectomorfo", "Orientado a fuerza explosiva.")
    ]
    for row_data in datos_2:
        row = tabla_2.add_row().cells
        for i, text in enumerate(row_data):
            row[i].text = text
            row[i].paragraphs[0].runs[0].font.size = Pt(8)
            if i == 0: row[i].paragraphs[0].runs[0].bold = True
            set_cell_border(row[i], top={"val": "nil"}, bottom={"val": "single", "sz": 4, "color": "E2E8F0"}, left={"val": "nil"}, right={"val": "nil"})

    doc.add_paragraph()
    
    # Gráfico Somatocarta
    p_soma = doc.add_paragraph()
    p_soma.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_soma.add_run().add_picture('somatocarta.png', width=Inches(4.0))

    # 5. TABLA 3: CARDIOMETABÓLICO
    doc.add_heading('3. Indicadores Cardiometabólicos', level=2).runs[0].font.color.rgb = RGBColor(15, 118, 110)
    tabla_3 = doc.add_table(rows=1, cols=4)
    tabla_3.style = 'Table Grid'
    hdr_3 = tabla_3.rows[0].cells
    cabeceras_3 = ["Indicador", "Resultado", "Referencia", "Observación Kinésica"]
    for i, text in enumerate(cabeceras_3):
        hdr_3[i].text = text
        hdr_3[i].paragraphs[0].runs[0].bold = True
        hdr_3[i].paragraphs[0].runs[0].font.size = Pt(8)
        hdr_3[i].paragraphs[0].runs[0].font.color.rgb = RGBColor(15, 118, 110)
        set_cell_border(hdr_3[i], bottom={"val": "single", "sz": 12, "color": "0F766E"}, top={"val": "nil"}, left={"val": "nil"}, right={"val": "nil"})

    datos_3 = [
        ("Índice Cintura / Cadera", "0.78", "< 0.83", "Riesgo cardiovascular bajo."),
        ("Suma 6 Pliegues", "35.9 mm", "< 45.0 mm", "Excelente perfil magro."),
        ("Metabolismo Basal", "1801 kcal", "Harris & Benedict", "Gasto mínimo en reposo."),
        ("Gasto Total (GET)", "3782 kcal", "Intensa (x2.1)", "Requerimiento para entrenamiento.")
    ]
    for row_data in datos_3:
        row = tabla_3.add_row().cells
        for i, text in enumerate(row_data):
            row[i].text = text
            row[i].paragraphs[0].runs[0].font.size = Pt(8)
            if i == 0: row[i].paragraphs[0].runs[0].bold = True
            set_cell_border(row[i], top={"val": "nil"}, bottom={"val": "single", "sz": 4, "color": "E2E8F0"}, left={"val": "nil"}, right={"val": "nil"})
    
    doc.add_paragraph()

    # 6. CONCLUSIÓN
    doc.add_heading('4. Conclusión y Recomendaciones', level=2).runs[0].font.color.rgb = RGBColor(15, 118, 110)
    p_interp = doc.add_paragraph(
        "El evaluado presenta una composición corporal altamente eficiente para el rendimiento deportivo, destacando un desarrollo muscular sobresaliente (57.2%) e índices de adiposidad mínimos (15.4%). "
        "Su perfil somatotípico se clasifica como Mesomorfo Puro (1.4-8.1-1.7).\n\n"
        "1. Mantenimiento Estructural: Preservar la masa muscular activa mediante sobrecarga progresiva.\n"
        "2. Soporte Energético: Ajustar la ingesta calórica al GET (~3782 kcal/día).\n"
        "3. Monitoreo: Reevaluar en 8 a 12 semanas."
    )
    p_interp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_interp.runs[0].font.size = Pt(9)

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
                label="📥 Descargar Dossier en Word",
                data=doc_bytes,
                file_name="Dossier_Kinesiologico_Completo.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
