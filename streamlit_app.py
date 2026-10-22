
import io
import re
import math
import hashlib
from datetime import datetime

import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, KeepTogether
)

st.set_page_config(
    page_title="Informes antropométricos",
    page_icon="📊",
    layout="wide"
)

ERRORS = [
    "#REF!", "#DIV/0!", "#VALUE!", "#NAME?",
    "#N/A", "#NUM!", "#NULL!"
]

st.title("Informes antropométricos")
st.caption("Auditoría de datos, indicadores básicos e informes profesionales")

st.warning(
    "Herramienta de apoyo profesional. Comprueba los datos, las unidades "
    "y las ecuaciones antes de entregar el informe. No sustituye la "
    "validación técnica ni la evaluación clínica."
)


def clean_value(value):
    if pd.isna(value):
        return ""
    if isinstance(value, (datetime, pd.Timestamp)):
        return value.strftime("%d-%m-%Y")
    if isinstance(value, float):
        if not math.isfinite(value):
            return str(value)
        return round(value, 4)
    return str(value).strip()


def excel_sheets(file_bytes):
    """Lee hojas y audita fórmulas en libros XLSX/XLSM."""
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    sheets = {}
    for name in xls.sheet_names:
        sheets[name] = pd.read_excel(
            io.BytesIO(file_bytes),
            sheet_name=name,
            header=None
        )
    return sheets


def audit_formulas(file_bytes, filename):
    findings = []
    formula_count = 0

    if not filename.lower().endswith((".xlsx", ".xlsm")):
        return formula_count, findings, (
            "La auditoría de fórmulas de Excel requiere .xlsx o .xlsm."
        )

    try:
        import openpyxl

        formulas_wb = openpyxl.load_workbook(
            io.BytesIO(file_bytes), data_only=False, read_only=True
        )
        values_wb = openpyxl.load_workbook(
            io.BytesIO(file_bytes), data_only=True, read_only=True
        )

        for sheet_name in formulas_wb.sheetnames:
            ws_formula = formulas_wb[sheet_name]
            ws_values = values_wb[sheet_name]

            for row in ws_formula.iter_rows():
                for cell in row:
                    value = cell.value
                    if isinstance(value, str) and value.startswith("="):
                        formula_count += 1
                        cached = ws_values[cell.coordinate].value
                        if isinstance(cached, str) and any(
                            err in cached.upper() for err in ERRORS
                        ):
                            findings.append({
                                "Hoja": sheet_name,
                                "Celda": cell.coordinate,
                                "Tipo": "Error en resultado almacenado",
                                "Detalle": cached,
                                "Fórmula": value
                            })
        formulas_wb.close()
        values_wb.close()

        note = (
            "Los resultados almacenados pueden estar ausentes o desactualizados "
            "si Excel no ha recalculado el libro. Esta revisión no ejecuta "
            "ni valida las fórmulas."
        )
        return formula_count, findings, note

    except Exception as e:
        return formula_count, findings, (
            "No se pudo completar la auditoría de fórmulas: " + str(e)
        )


def workbook_quality(sheets):
    findings = []
    total_rows = 0
    total_cols = 0
    total_missing = 0
    total_duplicates = 0

    for name, raw in sheets.items():
        df = raw.dropna(how="all").dropna(axis=1, how="all")
        total_rows += len(df)
        total_cols += len(df.columns)
        missing = int(df.isna().sum().sum())
        total_missing += missing

        # Sin encabezados confirmados, las filas duplicadas son solo
        # duplicados exactos de filas; no necesariamente errores.
        duplicates = int(df.duplicated().sum())
        total_duplicates += duplicates

        findings.append({
            "Hoja": name,
            "Filas con contenido": len(df),
            "Columnas con contenido": len(df.columns),
            "Celdas vacías": missing,
            "Filas exactamente duplicadas": duplicates
        })

    return {
        "hojas": len(sheets),
        "filas": total_rows,
        "columnas": total_cols,
        "faltantes": total_missing,
        "duplicados": total_duplicates,
        "detalle": findings
    }


def find_numeric_indicators(sheets):
    """
    Calcula IMC y relaciones cintura/estatura o cintura/cadera solo
    cuando encuentra columnas con nombres reconocibles y unidades
    compatibles. No infiere valores a partir de celdas desconocidas.
    """
    results = []
    notes = []

    aliases = {
        "peso_kg": ["peso kg", "peso (kg)", "weight kg", "weight (kg)"],
        "altura_cm": [
            "estatura cm", "altura cm", "talla cm",
            "estatura (cm)", "altura (cm)", "talla (cm)"
        ],
        "altura_m": [
            "estatura m", "altura m", "talla m",
            "estatura (m)", "altura (m)", "talla (m)"
        ],
        "cintura_cm": [
            "cintura cm", "perimetro cintura", "perímetro cintura",
            "waist cm", "cintura (cm)"
        ],
        "cadera_cm": [
            "cadera cm", "perimetro cadera", "perímetro cadera",
            "hip cm", "cadera (cm)"
        ]
    }

    for sheet_name, raw in sheets.items():
        # Se prueban encabezados solo en la primera fila.
        df = raw.copy()
        if df.empty:
            continue
        headers = [
            str(x).strip().lower() if pd.notna(x) else ""
            for x in df.iloc[0].tolist()
        ]
        if not any(headers):
            continue

        body = df.iloc[1:].copy()
        body.columns = headers

        found = {}
        for key, options in aliases.items():
            for col in body.columns:
                normalized = re.sub(r"\s+", " ", str(col)).strip()
                if normalized in options:
                    found[key] = col
                    break

        # Requiere que la información esté organizada en columnas
        # y que peso/estatura estén expresados en unidades reconocidas.
        if "peso_kg" in found and (
            "altura_cm" in found or "altura_m" in found
        ):
            for _, row in body.iterrows():
                try:
                    weight = pd.to_numeric(row[found["peso_kg"]], errors="coerce")
                    if "altura_cm" in found:
                        height = pd.to_numeric(
                            row[found["altura_cm"]], errors="coerce"
                        ) / 100
                    else:
                        height = pd.to_numeric(
                            row[found["altura_m"]], errors="coerce"
                        )

                    if pd.notna(weight) and pd.notna(height) and weight > 0 and height > 0:
                        bmi = float(weight) / float(height) ** 2
                        results.append({
                            "Hoja": sheet_name,
                            "Indicador": "IMC",
                            "Resultado": round(bmi, 2),
                            "Unidades": "kg/m²",
                            "Método": "Peso (kg) / estatura (m)²"
                        })
                except (TypeError, ValueError, ZeroDivisionError):
                    pass

        if "cintura_cm" in found and "altura_cm" in found:
            for _, row in body.iterrows():
                try:
                    waist = pd.to_numeric(row[found["cintura_cm"]], errors="coerce")
                    height = pd.to_numeric(row[found["altura_cm"]], errors="coerce")
                    if pd.notna(waist) and pd.notna(height) and height > 0:
                        results.append({
                            "Hoja": sheet_name,
                            "Indicador": "Cintura/estatura",
                            "Resultado": round(float(waist) / float(height), 3),
                            "Unidades": "adimensional",
                            "Método": "Cintura (cm) / estatura (cm)"
                        })
                except (TypeError, ValueError, ZeroDivisionError):
                    pass

        if "cintura_cm" in found and "cadera_cm" in found:
            for _, row in body.iterrows():
                try:
                    waist = pd.to_numeric(row[found["cintura_cm"]], errors="coerce")
                    hip = pd.to_numeric(row[found["cadera_cm"]], errors="coerce")
                    if pd.notna(waist) and pd.notna(hip) and hip > 0:
                        results.append({
                            "Hoja": sheet_name,
                            "Indicador": "Cintura/cadera",
                            "Resultado": round(float(waist) / float(hip), 3),
                            "Unidades": "adimensional",
                            "Método": "Cintura (cm) / cadera (cm)"
                        })
                except (TypeError, ValueError, ZeroDivisionError):
                    pass

    if not results:
        notes.append(
            "No se calcularon indicadores automáticamente. El archivo puede "
            "usar otro diseño, nombres de variables distintos, medidas "
            "dispuestas en filas o unidades no identificadas. No se asumieron "
            "equivalencias de unidades."
        )

    return results, notes


def dataframe_to_records(df, limit=100):
    data = df.head(limit).copy()
    data = data.fillna("")
    return [[str(c) for c in data.columns]] + [
        [str(clean_value(v)) for v in row]
        for row in data.itertuples(index=False, name=None)
    ]


def build_docx(filename, sheets, quality, formula_count,
                formula_errors, formula_note, indicators, comments):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)

    title = doc.add_heading("INFORME DE REVISIÓN ANTROPOMÉTRICA", 0)
    title.alignment = 1
    p = doc.add_paragraph(
        "Auditoría del archivo, calidad de datos e indicadores básicos"
    )
    p.alignment = 1

    doc.add_paragraph(
        f"Archivo analizado: {filename}\n"
        f"Fecha de generación: {datetime.now().strftime('%d-%m-%Y %H:%M')}"
    )

    doc.add_heading("1. Resumen ejecutivo", level=1)
    doc.add_paragraph(
        f"Se revisaron {quality['hojas']} hoja(s), con "
        f"{quality['filas']} filas y {quality['columnas']} columnas con contenido. "
        f"Se identificaron {quality['faltantes']} celdas vacías en el conjunto "
        f"de datos revisado y {quality['duplicados']} filas exactamente duplicadas. "
        f"Se detectaron {len(formula_errors)} resultados de fórmula almacenados "
        "con posibles errores visibles."
    )

    doc.add_paragraph(
        "Estos recuentos describen la estructura del archivo; no demuestran "
        "por sí solos que los datos sean válidos o inválidos. Las filas "
        "duplicadas y celdas vacías deben revisarse según el contexto."
    )

    doc.add_heading("2. Calidad y estructura del archivo", level=1)
    table = doc.add_table(rows=1, cols=5)
    table.style = "Light Shading Accent 1"
    for cell, text in zip(
        table.rows[0].cells,
        ["Hoja", "Filas", "Columnas", "Vacías", "Duplicadas"]
    ):
        cell.text = text

    for r in quality["detalle"]:
        cells = table.add_row().cells
        vals = [
            r["Hoja"], r["Filas con contenido"], r["Columnas con contenido"],
            r["Celdas vacías"], r["Filas exactamente duplicadas"]
        ]
        for cell, val in zip(cells, vals):
            cell.text = str(val)

    doc.add_heading("3. Auditoría de fórmulas", level=1)
    doc.add_paragraph(f"Fórmulas encontradas: {formula_count}.")
    doc.add_paragraph(formula_note)

    if formula_errors:
        t = doc.add_table(rows=1, cols=4)
        t.style = "Light Shading Accent 1"
        for c, text in zip(
            t.rows[0].cells, ["Hoja", "Celda", "Resultado", "Fórmula"]
        ):
            c.text = text
        for r in formula_errors:
            cells = t.add_row().cells
            vals = [r["Hoja"], r["Celda"], r["Detalle"], r["Fórmula"]]
            for cell, val in zip(cells, vals):
                cell.text = str(val)
    else:
        doc.add_paragraph(
            "No se encontraron errores de resultado almacenado durante la "
            "auditoría, o no fue posible acceder a dichos resultados. Esto "
            "no demuestra que todas las fórmulas sean correctas."
        )

    doc.add_heading("4. Indicadores básicos calculados", level=1)
    if indicators:
        t = doc.add_table(rows=1, cols=4)
        t.style = "Light Shading Accent 1"
        for c, text in zip(
            t.rows[0].cells, ["Hoja", "Indicador", "Resultado", "Método"]
        ):
            c.text = text
        for r in indicators:
            cells = t.add_row().cells
            vals = [r["Hoja"], r["Indicador"], r["Resultado"], r["Método"]]
            for cell, val in zip(cells, vals):
                cell.text = str(val)
    else:
        doc.add_paragraph(
            "No se pudieron calcular indicadores automáticamente con "
            "suficiente certeza a partir de los nombres de columnas y unidades."
        )

    doc.add_paragraph(
        "El IMC es un indicador antropométrico general y no mide directamente "
        "la adiposidad ni la composición corporal. Las relaciones cintura/"
        "estatura y cintura/cadera deben interpretarse considerando población, "
        "edad, sexo y contexto clínico o deportivo. No se asignaron categorías "
        "automáticas."
    )

    doc.add_heading("5. Observaciones del profesional", level=1)
    doc.add_paragraph(comments or "No se añadieron observaciones.")
    doc.add_heading("6. Limitaciones", level=1)
    doc.add_paragraph(
        "Este informe realiza una revisión automatizada de la estructura y "
        "de indicadores básicos reconocibles. No valida protocolos de medición, "
        "calibración, ecuaciones de composición corporal, referencias "
        "normativas ni diagnósticos. Las ecuaciones deben verificarse frente "
        "al protocolo y la bibliografía adoptados por el profesional."
    )

    doc.add_heading("Anexo. Vista de los datos originales", level=1)
    for sheet_name, raw in sheets.items():
        doc.add_heading(f"Hoja: {sheet_name}", level=2)
        preview = raw.dropna(how="all").dropna(axis=1, how="all").head(25)
        if preview.empty:
            doc.add_paragraph("Sin datos visibles.")
            continue
        t = doc.add_table(rows=0, cols=max(1, len(preview.columns)))
        t.style = "Table Grid"
        for _, row in preview.iterrows():
            cells = t.add_row().cells
            for i, cell in enumerate(cells):
                cell.text = str(clean_value(row.iloc[i]))
        doc.add_paragraph(
            "Se muestran como máximo 25 filas de esta hoja en el anexo."
        )

    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()


def build_pdf(filename, quality, formula_count, formula_errors,
              formula_note, indicators, comments):
    output = io.BytesIO()
    pdf = SimpleDocTemplate(
        output, pagesize=A4,
        rightMargin=1.6 * cm, leftMargin=1.6 * cm,
        topMargin=1.6 * cm, bottomMargin=1.6 * cm
    )

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="TitleCentered", parent=styles["Title"],
        alignment=TA_CENTER, textColor=colors.HexColor("#17365D"),
        fontSize=18, leading=22, spaceAfter=10
    ))
    styles.add(ParagraphStyle(
        name="SmallCell", parent=styles["BodyText"],
        fontSize=7, leading=9
    ))
    story = []

    def para(text, style="BodyText"):
        safe = (
            str(text).replace("&", "&amp;")
            .replace("<", "&lt;").replace(">", "&gt;")
        )
        story.append(Paragraph(safe, styles[style]))
        story.append(Spacer(1, 5))

    def table_from_rows(rows, widths=None):
        formatted = [
            [Paragraph(str(v).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"),
                       styles["SmallCell"]) for v in row]
            for row in rows
        ]
        t = Table(formatted, colWidths=widths, repeatRows=1, hAlign="LEFT")
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9EAF7")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#17365D")),
            ("GRID", (0, 0), (-1, -1), 0.35, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(t)
        story.append(Spacer(1, 10))

    para("INFORME DE REVISIÓN ANTROPOMÉTRICA", "TitleCentered")
    para("Auditoría del archivo, calidad de datos e indicadores básicos")
    para(f"Archivo analizado: {filename}")
    para(f"Generado: {datetime.now().strftime('%d-%m-%Y %H:%M')}")

    para("1. Resumen ejecutivo", "Heading1")
    para(
        f"Hojas: {quality['hojas']}; filas: {quality['filas']}; "
        f"columnas: {quality['columnas']}; celdas vacías: {quality['faltantes']}; "
        f"filas exactamente duplicadas: {quality['duplicados']}; "
        f"resultados de fórmula con errores visibles: {len(formula_errors)}."
    )
    para(
        "Los recuentos describen la estructura del archivo. Las filas "
        "duplicadas o las celdas vacías no constituyen necesariamente errores."
    )

    para("2. Calidad y estructura", "Heading1")
    rows = [["Hoja", "Filas", "Columnas", "Vacías", "Duplicadas"]]
    for r in quality["detalle"]:
        rows.append([
            r["Hoja"], str(r["Filas con contenido"]),
            str(r["Columnas con contenido"]), str(r["Celdas vacías"]),
            str(r["Filas exactamente duplicadas"])
        ])
    table_from_rows(rows)

    para("3. Auditoría de fórmulas", "Heading1")
    para(f"Fórmulas encontradas: {formula_count}.")
    para(formula_note)
    if formula_errors:
        rows = [["Hoja", "Celda", "Resultado", "Fórmula"]]
        rows.extend([
            [r["Hoja"], r["Celda"], r["Detalle"], r["Fórmula"]]
            for r in formula_errors
        ])
        table_from_rows(rows)
    else:
        para("No se detectaron errores almacenados visibles en esta revisión.")

    para("4. Indicadores básicos", "Heading1")
    if indicators:
        rows = [["Hoja", "Indicador", "Resultado", "Método"]]
        rows.extend([
            [r["Hoja"], r["Indicador"], str(r["Resultado"]), r["Método"]]
            for r in indicators
        ])
        table_from_rows(rows)
    else:
        para(
            "No se calcularon indicadores automáticamente: los nombres de "
            "columnas o las unidades no pudieron reconocerse con seguridad."
        )

    para(
        "El IMC no mide directamente la adiposidad ni la composición corporal. "
        "No se asignaron categorías clínicas o deportivas automáticas."
    )
    para("5. Observaciones del profesional", "Heading1")
    para(comments or "No se añadieron observaciones.")
    para("6. Limitaciones", "Heading1")
    para(
        "La herramienta no valida protocolos, calibración, ecuaciones de "
        "composición corporal ni referencias normativas. La interpretación "
        "final requiere revisión profesional."
    )

    pdf.build(story)
    return output.getvalue()


uploaded = st.file_uploader(
    "Carga tu archivo de evaluación",
    type=["xlsx", "xlsm", "xls"],
    help="Usa una copia sin información identificable mientras pruebas la aplicación."
)

comments = st.text_area(
    "Observaciones profesionales (opcional)",
    placeholder="Contexto de la evaluación, limitaciones o comentarios..."
)

if uploaded is not None:
    file_bytes = uploaded.getvalue()
    file_hash = hashlib.sha256(file_bytes).hexdigest()

    try:
        sheets = excel_sheets(file_bytes)
        quality = workbook_quality(sheets)
        formula_count, formula_errors, formula_note = audit_formulas(
            file_bytes, uploaded.name
        )
        indicators, indicator_notes = find_numeric_indicators(sheets)

        st.success("Excel leído. Revisa los resultados antes de generar el informe.")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Hojas", quality["hojas"])
        c2.metric("Filas con contenido", quality["filas"])
        c3.metric("Celdas vacías", quality["faltantes"])
        c4.metric("Errores visibles de fórmulas", len(formula_errors))

        with st.expander("Revisar hojas y datos", expanded=True):
            for name, raw in sheets.items():
                st.markdown(f"**Hoja: {name}**")
                st.dataframe(raw, use_container_width=True, height=250)

        st.subheader("Calidad de datos")
        st.dataframe(pd.DataFrame(quality["detalle"]), use_container_width=True)

        st.subheader("Auditoría de fórmulas")
        st.write(f"Fórmulas encontradas: **{formula_count}**")
        st.caption(formula_note)
        if formula_errors:
            st.dataframe(pd.DataFrame(formula_errors), use_container_width=True)
        else:
            st.info(
                "No se encontraron resultados almacenados con errores visibles. "
                "Esto no certifica que las fórmulas sean correctas."
            )

        st.subheader("Indicadores básicos")
        if indicators:
            st.dataframe(pd.DataFrame(indicators), use_container_width=True)
        else:
            st.info(
                "No se calcularon indicadores automáticamente. Es posible que "
                "las medidas estén organizadas en filas, tengan encabezados "
                "distintos o no indiquen unidades compatibles."
            )
        for note in indicator_notes:
            st.caption(note)

        st.caption(f"Identificador del archivo (SHA-256): {file_hash[:16]}…")

        if st.button("Generar informes Word y PDF", type="primary"):
            docx_bytes = build_docx(
                uploaded.name, sheets, quality, formula_count,
                formula_errors, formula_note, indicators, comments
            )
            pdf_bytes = build_pdf(
                uploaded.name, quality, formula_count,
                formula_errors, formula_note, indicators, comments
            )

            st.session_state["docx_bytes"] = docx_bytes
            st.session_state["pdf_bytes"] = pdf_bytes
            st.session_state["report_hash"] = file_hash

        if (
            "docx_bytes" in st.session_state
            and "pdf_bytes" in st.session_state
            and st.session_state.get("report_hash") == file_hash
        ):
            st.success("Informes generados.")
            st.download_button(
                "Descargar informe Word",
                data=st.session_state["docx_bytes"],
                file_name="informe_antropometrico.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            st.download_button(
                "Descargar informe PDF",
                data=st.session_state["pdf_bytes"],
                file_name="informe_antropometrico.pdf",
                mime="application/pdf"
            )

    except Exception as e:
        st.error("No se pudo analizar el archivo.")
        st.exception(e)
