from __future__ import annotations

import io
import re
from datetime import date, datetime
from typing import Any, Dict, Iterable, List, Optional, Tuple

import pandas as pd
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from src.config import (
    COLUMNAS_EVALUACIONES,
    COLUMNAS_GC72,
    COLUMNAS_SESIONES,
    LOGO_ICONTEC,
    LOGO_POLI,
    MAPA_TABLA_WORD_GC72,
    NUMERICAS_ENTERAS_GC72,
    NUMERICAS_PORCENTAJE_GC72,
    PREFORMAS_GC72,
    TEMPLATE_GC72,
)


def set_cell_width(cell, width_inches: float):
    cell.width = Inches(width_inches)
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_w = tc_pr.find(qn("w:tcW"))
    if tc_w is None:
        tc_w = OxmlElement("w:tcW")
        tc_pr.append(tc_w)
    tc_w.set(qn("w:w"), str(int(width_inches * 1440)))
    tc_w.set(qn("w:type"), "dxa")


def fijar_anchos_tabla(tabla, anchos: List[float]):
    tabla.autofit = False
    for row in tabla.rows:
        for idx, width in enumerate(anchos):
            if idx < len(row.cells):
                set_cell_width(row.cells[idx], width)


def shade_cell(cell, fill: str):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill.replace("#", ""))


def set_cell_text(
    cell,
    texto: str,
    font_size: float = 8.0,
    bold: bool = False,
    align=WD_ALIGN_PARAGRAPH.CENTER,
    color: Optional[str] = None,
):
    cell.text = ""
    parrafo = cell.paragraphs[0]
    parrafo.alignment = align
    parrafo.paragraph_format.space_after = Pt(0)
    run = parrafo.add_run(str(texto) if texto is not None else "")
    run.bold = bold
    run.font.size = Pt(font_size)
    if color:
        color = color.replace("#", "")
        run.font.color.rgb = RGBColor.from_string(color)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def add_paragraph_in_cell(cell, texto: str, font_size: float = 8.5, bold: bool = False, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    p = cell.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.05
    r = p.add_run(texto)
    r.bold = bold
    r.font.size = Pt(font_size)
    return p


def set_section_text(cell, titulo: str, texto: str, font_size: float = 8.5):
    cell.text = ""
    p1 = cell.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p1.add_run(titulo)
    r1.bold = True
    r1.font.size = Pt(9)
    if texto:
        for parte in str(texto).split("\n"):
            if parte.strip():
                add_paragraph_in_cell(cell, parte.strip(), font_size=font_size)


def set_label_value(cell, etiqueta: str, valor: str):
    cell.text = ""
    parrafo = cell.paragraphs[0]
    parrafo.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r1 = parrafo.add_run(etiqueta)
    r1.bold = True
    r1.font.size = Pt(10)
    r2 = parrafo.add_run(f" {valor}" if valor else "")
    r2.font.size = Pt(10)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def aplicar_bordes_tabla(tabla, color="000000", size="4"):
    tbl = tabla._tbl
    tbl_pr = tbl.tblPr
    borders = tbl_pr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = f"w:{edge}"
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), size)
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), color)


def remover_parrafo(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    paragraph._p = paragraph._element = None


def formato_entero(valor: Any) -> str:
    if valor is None or pd.isna(valor) or str(valor).strip() == "":
        return ""
    try:
        return str(int(round(float(valor))))
    except Exception:
        return str(valor)


def formato_porcentaje(valor: Any) -> str:
    if valor is None or pd.isna(valor) or str(valor).strip() == "":
        return ""
    try:
        val = float(valor)
        if val <= 1.0 and val > 0:
            val = val * 100
        return f"{val:.1f}%".replace(".0%", "%")
    except Exception:
        txt = str(valor).strip()
        return txt if txt.endswith("%") else f"{txt}%"


def porcentaje(numerador: Any, denominador: Any) -> str:
    try:
        num = float(numerador)
        den = float(denominador)
        if den == 0:
            return "0%"
        return f"{(num / den) * 100:.1f}%".replace(".0%", "%")
    except Exception:
        return "0%"


def limpiar_df(df: pd.DataFrame, columnas: List[str]) -> pd.DataFrame:
    if df is None:
        return pd.DataFrame(columns=columnas)
    res = df.copy()
    for col in columnas:
        if col not in res.columns:
            res[col] = ""
    return res[columnas]


def normalizar_dataframe_gc72(df: pd.DataFrame, calcular_porcentajes: bool = True) -> pd.DataFrame:
    df = limpiar_df(df, COLUMNAS_GC72)
    if calcular_porcentajes:
        for idx, row in df.iterrows():
            matriculados = row.get("Estudiantes matriculados", "")
            df.at[idx, "Desertaron %"] = porcentaje(row.get("Desertaron N°", ""), matriculados)
            df.at[idx, "Aprueban evaluación parcial %"] = porcentaje(row.get("Aprueban evaluación parcial N°", ""), matriculados)
            df.at[idx, "Reprueban evaluación parcial %"] = porcentaje(row.get("Reprueban evaluación parcial N°", ""), matriculados)
            df.at[idx, "Aprueban a la fecha %"] = porcentaje(row.get("Aprueban a la fecha N°", ""), matriculados)
            df.at[idx, "Reprueban a la fecha %"] = porcentaje(row.get("Reprueban a la fecha N°", ""), matriculados)
    return df


def texto_preformas(seleccionadas: Iterable[str], categoria: str, adicional: str = "") -> str:
    partes = [PREFORMAS_GC72[categoria][k] for k in seleccionadas if k in PREFORMAS_GC72[categoria]]
    adicional = (adicional or "").strip()
    if adicional:
        partes.append(adicional)
    return " ".join(partes).strip()


def curso_key(row: Any, idx: int) -> str:
    codigo = re.sub(r"\W+", "_", str(row.get("Código", "")).strip())
    grupo = re.sub(r"\W+", "_", str(row.get("Grupo", "")).strip())
    asignatura = re.sub(r"\W+", "_", str(row.get("Asignatura", "")).strip())[:40]
    return f"curso_{idx}_{codigo}_{grupo}_{asignatura}"


def agregar_parrafo_antes(parrafo_referencia: Any, texto: str = "", bold_prefix: Optional[str] = None, space_after: int = 3):
    nuevo = parrafo_referencia.insert_paragraph_before()
    nuevo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    nuevo.paragraph_format.space_after = Pt(space_after)
    nuevo.paragraph_format.line_spacing = 1.05
    if bold_prefix and texto.startswith(bold_prefix):
        r1 = nuevo.add_run(bold_prefix)
        r1.bold = True
        r1.font.size = Pt(10)
        r2 = nuevo.add_run(texto[len(bold_prefix):])
        r2.font.size = Pt(10)
    else:
        r = nuevo.add_run(texto)
        r.font.size = Pt(10)
    return nuevo


def construir_analisis(cursos: pd.DataFrame, analisis_por_curso: Dict[str, Dict[str, str]], modo: str) -> List[Dict[str, str]]:
    bloques = []
    if modo == "Consolidado institucional":
        positivos, inconvenientes, propuestas = [], [], []
        for data in analisis_por_curso.values():
            if data.get("positivos"):
                positivos.append(data["positivos"])
            if data.get("inconvenientes"):
                inconvenientes.append(data["inconvenientes"])
            if data.get("propuestas"):
                propuestas.append(data["propuestas"])
        bloques.append({
            "titulo": "Análisis consolidado de los cursos reportados",
            "positivos": " ".join(dict.fromkeys(positivos)),
            "inconvenientes": " ".join(dict.fromkeys(inconvenientes)),
            "propuestas": " ".join(dict.fromkeys(propuestas)),
        })
        return bloques

    for idx, row in cursos.iterrows():
        key = curso_key(row, idx)
        nombre = str(row.get("Asignatura", "")).strip() or "Curso sin nombre"
        codigo = str(row.get("Código", "")).strip()
        grupo = str(row.get("Grupo", "")).strip()
        detalles = []
        if codigo:
            detalles.append(f"Código {codigo}")
        if grupo:
            detalles.append(f"Grupo {grupo}")
        encabezado = f"{nombre} ({' - '.join(detalles)})" if detalles else nombre
        data = analisis_por_curso.get(key, {})
        bloques.append({
            "titulo": encabezado,
            "positivos": data.get("positivos", ""),
            "inconvenientes": data.get("inconvenientes", ""),
            "propuestas": data.get("propuestas", ""),
        })
    return bloques


def crear_informe_gc72_docx(docente: str, periodo: str, fecha_entrega: date | str, cursos: pd.DataFrame, bloques_analisis: List[Dict[str, str]]) -> bytes:
    if not TEMPLATE_GC72.exists():
        raise FileNotFoundError(f"No se encontró la plantilla: {TEMPLATE_GC72}")
    doc = Document(str(TEMPLATE_GC72))
    cursos_norm = normalizar_dataframe_gc72(cursos, calcular_porcentajes=False)

    datos = doc.tables[0]
    set_label_value(datos.rows[0].cells[0], "DOCENTE:", docente)
    set_label_value(datos.rows[1].cells[0], "PERÍODO ACADÉMICO:", periodo)
    fecha_texto = fecha_entrega.strftime("%d/%m/%Y") if hasattr(fecha_entrega, "strftime") else str(fecha_entrega)
    set_label_value(datos.rows[2].cells[0], "FECHA DE ENTREGA:", fecha_texto)

    tabla = doc.tables[1]
    fijar_anchos_tabla(tabla, [0.62, 0.55, 1.38, 0.72, 0.68, 0.90, 0.43, 0.38, 0.55, 0.38, 0.55, 0.40, 0.50, 0.38, 0.50, 0.38])
    fila_inicio = 3
    filas_necesarias = max(8, len(cursos_norm))
    while len(tabla.rows) < fila_inicio + filas_necesarias:
        tabla.add_row()

    for i in range(filas_necesarias):
        row_cells = tabla.rows[fila_inicio + i].cells
        if i < len(cursos_norm):
            registro = cursos_norm.iloc[i]
            for j, col in enumerate(MAPA_TABLA_WORD_GC72):
                valor = registro.get(col, "")
                if col in NUMERICAS_ENTERAS_GC72:
                    texto = formato_entero(valor)
                elif col in NUMERICAS_PORCENTAJE_GC72:
                    texto = formato_porcentaje(valor)
                    if j >= 7:
                        texto = texto.replace("%", "")
                else:
                    texto = "" if pd.isna(valor) else str(valor).strip()
                set_cell_text(row_cells[j], texto, font_size=6.8 if j >= 6 else 7.2)
        else:
            for j in range(min(len(row_cells), len(MAPA_TABLA_WORD_GC72))):
                set_cell_text(row_cells[j], "", font_size=7.5)

    firma = None
    for p in doc.paragraphs:
        if p.text.strip().startswith("Firma:"):
            firma = p
            break
    if firma is None:
        firma = doc.add_paragraph("Firma: ___________________")

    agregar_parrafo_antes(firma, "")
    for bloque in bloques_analisis:
        titulo = bloque.get("titulo", "Curso")
        p_titulo = firma.insert_paragraph_before()
        p_titulo.paragraph_format.space_before = Pt(4)
        p_titulo.paragraph_format.space_after = Pt(2)
        run = p_titulo.add_run(titulo)
        run.bold = True
        run.font.size = Pt(10)

        positivos = bloque.get("positivos", "").strip() or "Sin observaciones específicas para este apartado."
        inconvenientes = bloque.get("inconvenientes", "").strip() or "No se reportan inconvenientes relevantes para el periodo informado."
        propuestas = bloque.get("propuestas", "").strip() or "Mantener seguimiento y retroalimentación permanente para sostener el avance del curso."

        agregar_parrafo_antes(firma, f"1. Aspectos positivos: {positivos}", bold_prefix="1. Aspectos positivos:")
        agregar_parrafo_antes(firma, f"2. Inconvenientes presentados: {inconvenientes}", bold_prefix="2. Inconvenientes presentados:")
        agregar_parrafo_antes(firma, f"3. Propuestas metodológicas: {propuestas}", bold_prefix="3. Propuestas metodológicas:")

    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


def configurar_documento_gc71(doc: Document):
    section = doc.sections[0]
    section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(0.35)
    section.bottom_margin = Inches(0.35)
    section.left_margin = Inches(0.35)
    section.right_margin = Inches(0.35)
    styles = doc.styles
    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(8.5)


def agregar_tabla_header_gc71(doc: Document):
    table = doc.add_table(rows=2, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    aplicar_bordes_tabla(table)
    widths = [2.55, 3.9, 1.75]
    for row in table.rows:
        for i, w in enumerate(widths):
            set_cell_width(row.cells[i], w)
    logo_cell = table.cell(0, 0).merge(table.cell(1, 0))
    title_cell = table.cell(0, 1).merge(table.cell(1, 1))
    if LOGO_POLI.exists():
        logo_cell.text = ""
        p = logo_cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(str(LOGO_POLI), width=Inches(1.55))
    else:
        set_cell_text(logo_cell, "POLITÉCNICO COLOMBIANO\nJAIME ISAZA CADAVID", 8, True)
    set_cell_text(title_cell, "GUÍA DIDÁCTICA DE ASIGNATURA Y\nCONCERTACIÓN DE EVALUACIÓN", 12, True)
    set_cell_text(table.cell(0, 2), "Código: FD-GC71", 10, False)
    set_cell_text(table.cell(1, 2), "Versión: 09", 10, True)
    return table


def agregar_fila_seccion(table, titulo: str, cols: int, fill="FFF2CC"):
    row = table.add_row()
    cell = row.cells[0].merge(row.cells[cols - 1])
    shade_cell(cell, fill)
    set_cell_text(cell, titulo, 8.5, True)
    return row


def agregar_tabla_identificacion_gc71(doc: Document, datos: Dict[str, str]):
    doc.add_paragraph()
    t = doc.add_table(rows=0, cols=4)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    aplicar_bordes_tabla(t)
    widths = [2.2, 2.05, 1.85, 2.0]
    agregar_fila_seccion(t, "1. IDENTIFICACIÓN DE LA ASIGNATURA", 4)
    campos = [
        ("PROGRAMA ACADÉMICO", datos.get("programa", "")),
        ("ASIGNATURA", datos.get("asignatura", "")),
        ("CÓDIGO", datos.get("codigo", "")),
        ("ÁREA DE FORMACIÓN", datos.get("area", "")),
        ("PRERREQUISITO(S)", datos.get("prerrequisitos", "")),
        ("CORREQUISITO(S)", datos.get("correquisitos", "")),
    ]
    for etiqueta, valor in campos:
        row = t.add_row()
        c0 = row.cells[0].merge(row.cells[1])
        c1 = row.cells[2].merge(row.cells[3])
        set_cell_text(c0, etiqueta, 8.3, True, align=WD_ALIGN_PARAGRAPH.LEFT)
        set_cell_text(c1, valor, 8.3, False, align=WD_ALIGN_PARAGRAPH.LEFT)
    row = t.add_row()
    set_cell_text(row.cells[0], "TIPO DE ASIGNATURA", 8.3, True, align=WD_ALIGN_PARAGRAPH.LEFT)
    tipo = datos.get("tipo_asignatura", "Teórico-práctica")
    set_cell_text(row.cells[1], f"{'☒' if tipo == 'Teórica' else '☐'} Teórica", 8.3)
    set_cell_text(row.cells[2], f"{'☒' if tipo == 'Teórico-práctica' else '☐'} Teórico-práctica", 8.3)
    set_cell_text(row.cells[3], f"{'☒' if tipo == 'Práctica' else '☐'} Práctica", 8.3)
    row = t.add_row()
    set_cell_text(row.cells[0].merge(row.cells[1]), "NÚMERO DE CRÉDITOS", 8.3, True, align=WD_ALIGN_PARAGRAPH.LEFT)
    set_cell_text(row.cells[2].merge(row.cells[3]), datos.get("creditos", ""), 8.3, align=WD_ALIGN_PARAGRAPH.LEFT)
    row = t.add_row()
    set_cell_text(row.cells[0], "DISTRIBUCIÓN HORARIA SEMANAL", 8.3, True, align=WD_ALIGN_PARAGRAPH.LEFT)
    set_cell_text(row.cells[1], f"HTP: {datos.get('htp', '')}", 8.3)
    set_cell_text(row.cells[2], f"HTI: {datos.get('hti', '')}", 8.3)
    set_cell_text(row.cells[3], f"Total: {datos.get('ht_total', '')}", 8.3)
    for etiqueta, valor in [
        ("PROFESOR", datos.get("profesor", "")),
        ("CORREO ELECTRÓNICO", datos.get("correo", "")),
        ("GRUPO", datos.get("grupo", "")),
        ("PERÍODO ACADÉMICO", datos.get("periodo", "")),
    ]:
        row = t.add_row()
        c0 = row.cells[0].merge(row.cells[1])
        c1 = row.cells[2].merge(row.cells[3])
        set_cell_text(c0, etiqueta, 8.3, True, align=WD_ALIGN_PARAGRAPH.LEFT)
        set_cell_text(c1, valor, 8.3, align=WD_ALIGN_PARAGRAPH.LEFT)
    for row in t.rows:
        for i, w in enumerate(widths):
            if i < len(row.cells):
                set_cell_width(row.cells[i], w)
    return t


def agregar_seccion_texto_gc71(doc: Document, numero_titulo: str, texto: str):
    t = doc.add_table(rows=2, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    aplicar_bordes_tabla(t)
    shade_cell(t.rows[0].cells[0], "FFF2CC")
    set_cell_text(t.rows[0].cells[0], numero_titulo, 8.5, True)
    set_cell_text(t.rows[1].cells[0], texto, 8.2, align=WD_ALIGN_PARAGRAPH.LEFT)
    set_cell_width(t.rows[0].cells[0], 8.1)
    set_cell_width(t.rows[1].cells[0], 8.1)
    return t


def agregar_contenidos_gc71(doc: Document, sesiones_df: pd.DataFrame):
    t = doc.add_table(rows=0, cols=5)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    aplicar_bordes_tabla(t)
    widths = [0.75, 1.0, 2.45, 1.95, 1.95]
    agregar_fila_seccion(t, "7. CONTENIDOS TEMÁTICOS DE LA ASIGNATURA", 5)
    sesiones = limpiar_df(sesiones_df, COLUMNAS_SESIONES)
    for unidad, dfu in sesiones.groupby("Unidad", sort=False):
        agregar_fila_seccion(t, str(unidad).upper(), 5, fill="FFF2CC")
        row = t.add_row()
        headers = ["N° sesión", "Fecha", "Contenido por desarrollar", "Descripción del trabajo presencial", "Descripción trabajo independiente"]
        for i, h in enumerate(headers):
            shade_cell(row.cells[i], "F2F2F2")
            set_cell_text(row.cells[i], h, 7.5, True)
        for _, s in dfu.iterrows():
            row = t.add_row()
            fecha = s.get("Fecha")
            if hasattr(fecha, "strftime"):
                fecha_txt = fecha.strftime("%d/%m/%Y")
            else:
                fecha_txt = str(fecha)
            values = [
                s.get("N° sesión", ""),
                fecha_txt,
                s.get("Contenido por desarrollar", ""),
                s.get("Descripción del trabajo presencial", ""),
                s.get("Descripción trabajo independiente", ""),
            ]
            for i, v in enumerate(values):
                set_cell_text(row.cells[i], v, 7.2 if i >= 2 else 7.0, align=WD_ALIGN_PARAGRAPH.LEFT if i >= 2 else WD_ALIGN_PARAGRAPH.CENTER)
    for row in t.rows:
        for i, w in enumerate(widths):
            if i < len(row.cells):
                set_cell_width(row.cells[i], w)
    return t


def agregar_evaluaciones_gc71(doc: Document, asignatura: str, grupo: str, evaluaciones_df: pd.DataFrame):
    t = doc.add_table(rows=0, cols=4)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    aplicar_bordes_tabla(t)
    widths = [1.45, 4.15, 0.9, 1.6]
    agregar_fila_seccion(t, "11. EVALUACIÓN DE LA ASIGNATURA", 4)
    row = t.add_row()
    c0 = row.cells[0].merge(row.cells[1])
    c1 = row.cells[2].merge(row.cells[3])
    set_cell_text(c0, f"ASIGNATURA: {asignatura}", 7.6, True, align=WD_ALIGN_PARAGRAPH.LEFT)
    set_cell_text(c1, f"GRUPO: {grupo}", 7.6, True, align=WD_ALIGN_PARAGRAPH.LEFT)
    row = t.add_row()
    headers = ["TIPO DE EVALUACIÓN°", "PROCEDIMIENTO DE EVALUACIÓN\n(Descripción de la actividad evaluativa)", "VALOR (%)", "FECHA DE\nREALIZACIÓN"]
    for i, h in enumerate(headers):
        shade_cell(row.cells[i], "FFF2CC")
        set_cell_text(row.cells[i], h, 7.2, True)
    evaluaciones = limpiar_df(evaluaciones_df, COLUMNAS_EVALUACIONES)
    for _, ev in evaluaciones.iterrows():
        row = t.add_row()
        fecha = ev.get("Fecha de realización")
        if hasattr(fecha, "strftime"):
            fecha_txt = fecha.strftime("%d/%m/%Y")
        else:
            fecha_txt = str(fecha)
        vals = [ev.get("Tipo de evaluación", ""), ev.get("Procedimiento de evaluación", ""), ev.get("Valor (%)", ""), fecha_txt]
        for i, v in enumerate(vals):
            set_cell_text(row.cells[i], v, 7.2, align=WD_ALIGN_PARAGRAPH.LEFT if i in [0, 1] else WD_ALIGN_PARAGRAPH.CENTER)
    for row in t.rows:
        for i, w in enumerate(widths):
            if i < len(row.cells):
                set_cell_width(row.cells[i], w)
    return t


def agregar_evidencia_y_control_gc71(doc: Document, datos: Dict[str, str], representantes_df: pd.DataFrame):
    t = doc.add_table(rows=0, cols=3)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    aplicar_bordes_tabla(t)
    agregar_fila_seccion(t, "12. EVIDENCIA DE PRESENTACIÓN DE LA GUÍA Y CONCERTACIÓN DE EVALUACIÓN AL GRUPO DE ESTUDIANTES", 3)
    row = t.add_row()
    cell = row.cells[0].merge(row.cells[2])
    set_cell_text(cell, "Se deja constancia de socialización de la Guía Didáctica de Asignatura y aprobación de la concertación de evaluación según el reglamento estudiantil; para ello firman tres estudiantes en representación del grupo:", 7.5, align=WD_ALIGN_PARAGRAPH.CENTER)
    row = t.add_row()
    for i, h in enumerate(["Nombre de los estudiantes", "N° de cédula o carné estudiantil", "Firma"]):
        shade_cell(row.cells[i], "FFF2CC")
        set_cell_text(row.cells[i], h, 7.5, True)
    reps = representantes_df.copy() if representantes_df is not None else pd.DataFrame()
    for i in range(3):
        row = t.add_row()
        nombre = reps.iloc[i].get("Nombre", "") if i < len(reps) else ""
        docu = reps.iloc[i].get("Documento", "") if i < len(reps) else ""
        set_cell_text(row.cells[0], nombre, 7.5, align=WD_ALIGN_PARAGRAPH.LEFT)
        set_cell_text(row.cells[1], docu, 7.5)
        set_cell_text(row.cells[2], "", 7.5)
    row = t.add_row()
    for i, h in enumerate(["Nombre del docente del curso", "Cédula", "Firma"]):
        shade_cell(row.cells[i], "FFF2CC")
        set_cell_text(row.cells[i], h, 7.5, True)
    row = t.add_row()
    set_cell_text(row.cells[0], datos.get("profesor", ""), 7.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    set_cell_text(row.cells[1], datos.get("cedula_docente", ""), 7.5)
    set_cell_text(row.cells[2], "", 7.5)
    row = t.add_row()
    cell = row.cells[0].merge(row.cells[2])
    set_cell_text(cell, f"Fecha de socialización de la Guía Didáctica: {datos.get('fecha_socializacion', '')}", 7.5, True, align=WD_ALIGN_PARAGRAPH.LEFT)
    row = t.add_row()
    cell = row.cells[0].merge(row.cells[2])
    set_cell_text(cell, "Nota: El docente se compromete a devolver las evaluaciones, socializar la calificación con los estudiantes y a ingresar dicha calificación al sistema académico, correcta y oportunamente.", 7.2, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    for row in t.rows:
        set_cell_width(row.cells[0], 3.4)
        set_cell_width(row.cells[1], 2.0)
        set_cell_width(row.cells[2], 2.7)

    doc.add_paragraph()
    c = doc.add_table(rows=3, cols=2)
    c.alignment = WD_TABLE_ALIGNMENT.CENTER
    aplicar_bordes_tabla(c)
    cell = c.rows[0].cells[0].merge(c.rows[0].cells[1])
    shade_cell(cell, "FFF2CC")
    set_cell_text(cell, "CONTROL DE CAMBIOS Y VIGENCIA (DILIGENCIAR LOS DATOS ESPECÍFICOS)", 8, True)
    set_cell_text(c.rows[1].cells[0], "Fecha de Revisión por parte del Coordinador de Área:", 7.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    set_cell_text(c.rows[1].cells[1], datos.get("fecha_revision", ""), 7.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    set_cell_text(c.rows[2].cells[0], "Fecha de aprobación y acta de sesión del Comité de currículo del programa:", 7.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    set_cell_text(c.rows[2].cells[1], datos.get("fecha_aprobacion", ""), 7.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    for row in c.rows:
        set_cell_width(row.cells[0], 4.0)
        set_cell_width(row.cells[1], 4.1)


def crear_gc71_docx(
    datos: Dict[str, str],
    sesiones_df: pd.DataFrame,
    evaluaciones_df: pd.DataFrame,
    representantes_df: Optional[pd.DataFrame] = None,
) -> bytes:
    doc = Document()
    configurar_documento_gc71(doc)
    agregar_tabla_header_gc71(doc)
    agregar_tabla_identificacion_gc71(doc, datos)
    agregar_seccion_texto_gc71(doc, "2. JUSTIFICACIÓN", datos.get("justificacion", ""))
    agregar_seccion_texto_gc71(doc, "3. COMPETENCIAS A LAS QUE LE TRIBUTA LA ASIGNATURA", datos.get("competencias", ""))
    agregar_seccion_texto_gc71(doc, "4. RESULTADOS DE APRENDIZAJE A LOS QUE LE TRIBUTA LA ASIGNATURA", datos.get("resultados", ""))
    agregar_seccion_texto_gc71(doc, "5. OBJETIVOS DE APRENDIZAJE DE LA ASIGNATURA", f"OBJETIVO(S) GENERAL(ES)\n{datos.get('objetivo_general', '')}\n\nOBJETIVOS ESPECÍFICOS\n{datos.get('objetivos_especificos', '')}")
    agregar_seccion_texto_gc71(doc, "6. METODOLOGÍAS Y ESTRATEGIAS DIDÁCTICAS DE LA ASIGNATURA", datos.get("metodologias", ""))
    agregar_contenidos_gc71(doc, sesiones_df)
    agregar_seccion_texto_gc71(doc, "8. AMBIENTES DE APRENDIZAJE DE LA ASIGNATURA", datos.get("ambientes", ""))
    agregar_seccion_texto_gc71(doc, "9. MEDIOS EDUCATIVOS PARA LA ASIGNATURA", datos.get("medios", ""))
    agregar_seccion_texto_gc71(doc, "10. REFERENCIAS BIBLIOGRÁFICAS", datos.get("referencias", ""))
    agregar_evaluaciones_gc71(doc, datos.get("asignatura", ""), datos.get("grupo", ""), evaluaciones_df)
    agregar_evidencia_y_control_gc71(doc, datos, representantes_df if representantes_df is not None else pd.DataFrame())

    if LOGO_ICONTEC.exists():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run().add_picture(str(LOGO_ICONTEC), width=Inches(1.0))
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


def crear_informe_ejecutivo_institucional_docx(matriz: pd.DataFrame, resumen: Dict[str, Any]) -> bytes:
    doc = Document()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("INFORME EJECUTIVO INSTITUCIONAL DE GOBIERNO ACADÉMICO")
    r.bold = True
    r.font.size = Pt(14)
    doc.add_paragraph(f"Fecha de emisión: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    doc.add_paragraph(f"Total expedientes en sistema: {resumen.get('total', 0)}")
    doc.add_paragraph(f"Expedientes con aprobación bloqueante: {resumen.get('bloqueados', 0)}")
    doc.add_paragraph(f"Expedientes cerrados en norma: {resumen.get('cerrados', 0)}")

    if not matriz.empty:
        doc.add_heading("Matriz Consolidada de Cursos", level=2)
        table = doc.add_table(rows=1, cols=min(6, len(matriz.columns)))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        aplicar_bordes_tabla(table)
        cols = matriz.columns[:6]
        for i, c in enumerate(cols):
            set_cell_text(table.rows[0].cells[i], str(c), 8, True)
        for _, row in matriz.iterrows():
            r_cells = table.add_row().cells
            for i, c in enumerate(cols):
                set_cell_text(r_cells[i], str(row.get(c, "")), 7.5)

    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()
