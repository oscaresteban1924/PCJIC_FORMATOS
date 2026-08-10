from __future__ import annotations

import io
import os
import re
from datetime import date
from typing import Any, Dict, Iterable, List, Optional, Tuple

import docx
import pandas as pd
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor

from src.config import (
    COLUMNAS_EVALUACIONES,
    COLUMNAS_GC72,
    COLUMNAS_HORARIOS,
    COLUMNAS_MODULOS,
    COLUMNAS_SESIONES,
    MAPA_TABLA_WORD_GC72,
    NUMERICAS_ENTERAS_GC72,
    NUMERICAS_PORCENTAJE_GC72,
    PREFORMAS_GC72,
    TEMPLATE_GC71,
    TEMPLATE_GC72,
)


def porcentaje(numerador: Any, denominador: Any) -> str:
    try:
        num = float(numerador)
        den = float(denominador)
        if den == 0:
            return "0%"
        return f"{(num / den) * 100:.1f}%".replace(".0%", "%")
    except Exception:
        return "0%"


def formato_entero(valor: Any) -> str:
    if valor is None or pd.isna(valor) or str(valor).strip() == "":
        return "0"
    try:
        return str(int(round(float(valor))))
    except Exception:
        return str(valor)


def formato_porcentaje(valor: Any) -> str:
    if valor is None or pd.isna(valor) or str(valor).strip() == "":
        return "0%"
    txt = str(valor).strip()
    if txt.endswith("%"):
        return txt
    try:
        val = float(txt)
        if val <= 1.0 and val > 0:
            val = val * 100
        return f"{val:.1f}%".replace(".0%", "%")
    except Exception:
        return txt


def shade_cell(cell: Any, color_hex: str):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)


def set_cell_width(cell: Any, width_in_inches: float):
    cell.width = Inches(width_in_inches)
    tcPr = cell._tc.get_or_add_tcPr()
    tcW = parse_xml(f'<w:tcW {nsdecls("w")} w:w="{int(width_in_inches * 1440)}" w:type="dxa"/>')
    tcPr.append(tcW)


def fijar_anchos_tabla(table: Any, col_widths_in_inches: List[float]):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for row in table.rows:
        for idx, width in enumerate(col_widths_in_inches):
            if idx < len(row.cells):
                set_cell_width(row.cells[idx], width)


def aplicar_bordes_tabla(table: Any, color_hex: str = "D3D3D3"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="4" w:space="0" w:color="{color_hex}"/>'
        f'<w:left w:val="none"/>'
        f'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="002060"/>'
        f'<w:right w:val="none"/>'
        f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="{color_hex}"/>'
        f'<w:insideV w:val="none"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)


def set_cell_text(cell: Any, text: str, bold: bool = False, italic: bool = False, font_size_pt: float = 9.5, color_rgb: Tuple[int, int, int] = (0, 0, 0), align: WD_ALIGN_PARAGRAPH = WD_ALIGN_PARAGRAPH.LEFT):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.05
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(font_size_pt)
    run.font.color.rgb = RGBColor(*color_rgb)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def add_paragraph_in_cell(cell: Any, text: str = "", bold: bool = False, italic: bool = False, font_size_pt: float = 9.5, color_rgb: Tuple[int, int, int] = (0, 0, 0), space_after: float = 2):
    p = cell.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.05
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(font_size_pt)
    run.font.color.rgb = RGBColor(*color_rgb)
    return p


def remover_parrafo(p: Any):
    p._element.getparent().remove(p._element)


def set_label_value(cell: Any, label: str, value: str, font_size_pt: float = 9.5):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.05

    run_l = p.add_run(f"{label}: ")
    run_l.bold = True
    run_l.font.size = Pt(font_size_pt)
    run_l.font.color.rgb = RGBColor(0, 32, 96)

    run_v = p.add_run(str(value or ""))
    run_v.bold = False
    run_v.font.size = Pt(font_size_pt)
    run_v.font.color.rgb = RGBColor(30, 30, 30)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def set_section_text(cell: Any, text: str, font_size_pt: float = 9.5):
    cell.text = ""
    lineas = str(text or "").split("\n")
    first = True
    for line in lineas:
        txt = line.strip()
        if not txt:
            continue
        p = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.05
        run = p.add_run(txt)
        run.font.size = Pt(font_size_pt)
        run.font.color.rgb = RGBColor(30, 30, 30)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP


def normalizar_dataframe_gc72(df_input: pd.DataFrame, calcular_porcentajes: bool = True) -> pd.DataFrame:
    df = df_input.copy()
    for col in COLUMNAS_GC72:
        if col not in df.columns:
            df[col] = ""

    for idx, row in df.iterrows():
        asig = str(row.get("Asignatura", "")).strip()
        cod = str(row.get("Código", "")).strip()
        grp = str(row.get("Grupo", "")).strip()

        if not asig and not cod and not grp:
            continue

        matriculados = row.get("Estudiantes matriculados", 0)
        try:
            matriculados = int(round(float(matriculados))) if matriculados != "" else 0
        except Exception:
            matriculados = 0

        for col in NUMERICAS_ENTERAS_GC72:
            val = row.get(col, 0)
            try:
                df.at[idx, col] = int(round(float(val))) if val != "" else 0
            except Exception:
                df.at[idx, col] = 0

        if calcular_porcentajes:
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
        data = analisis_por_curso.get(key, {})
        bloques.append({
            "titulo": f"{nombre} | Código: {codigo} | Grupo: {grupo}",
            "positivos": data.get("positivos", ""),
            "inconvenientes": data.get("inconvenientes", ""),
            "propuestas": data.get("propuestas", ""),
        })
    return bloques


def crear_gc71_docx(
    datos: Dict[str, str],
    sesiones_df: pd.DataFrame,
    evaluaciones_df: pd.DataFrame,
    representantes_df: Optional[pd.DataFrame] = None,
) -> bytes:
    doc = docx.Document(TEMPLATE_GC71 if os.path.exists(TEMPLATE_GC71) else None)
    # Rellenado básico del documento Word
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()


def crear_informe_gc72_docx(
    docente: str,
    periodo: str,
    fecha_entrega: date | str,
    cursos_df: pd.DataFrame,
    bloques_analisis: List[Dict[str, str]],
) -> bytes:
    doc = docx.Document(TEMPLATE_GC72 if os.path.exists(TEMPLATE_GC72) else None)
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()


def crear_informe_ejecutivo_institucional_docx(matriz_df: pd.DataFrame, resumen: Dict[str, Any]) -> bytes:
    doc = docx.Document()
    doc.add_heading("POLITÉCNICO COLOMBIANO JAIME ISAZA CADAVID", level=1)
    doc.add_heading("Informe Ejecutivo de Gobierno Académico", level=2)
    p = doc.add_paragraph()
    p.add_run(f"Fecha de emisión: {date.today().strftime('%d/%m/%Y')}\n").bold = True
    p.add_run(f"Total cursos auditados: {resumen.get('total', 0)}\n")

    table = doc.add_table(rows=1, cols=6)
    hdr_cells = table.rows[0].cells
    hdr_titles = ["Código", "Grupo", "Asignatura", "Profesor", "Estado", "Calidad"]
    for i, title in enumerate(hdr_titles):
        hdr_cells[i].text = title
        shade_cell(hdr_cells[i], "002060")

    if matriz_df is not None and not matriz_df.empty:
        for _, row in matriz_df.iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text = str(row.get("Código", ""))
            row_cells[1].text = str(row.get("Grupo", ""))
            row_cells[2].text = str(row.get("Asignatura", ""))
            row_cells[3].text = str(row.get("Profesor", ""))
            row_cells[4].text = str(row.get("Estado", ""))
            row_cells[5].text = str(row.get("Score Calidad", ""))

    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()


def crear_gc71_html_imprimible(
    datos: Dict[str, str],
    sesiones_df: pd.DataFrame,
    evaluaciones_df: pd.DataFrame,
) -> str:
    """Genera una plantilla HTML formal lista para imprimir en PDF con CSS de alta fidelidad."""
    sesiones_rows = ""
    if sesiones_df is not None and not sesiones_df.empty:
        for _, r in sesiones_df.iterrows():
            sesiones_rows += f"""
            <tr>
                <td>{r.get('Unidad','')}</td>
                <td style="text-align:center;">{r.get('N° sesión','')}</td>
                <td style="text-align:center;">{r.get('Fecha','')}</td>
                <td style="text-align:center;">{r.get('Horario','')}</td>
                <td>{r.get('Contenido por desarrollar','')}</td>
                <td>{r.get('Descripción del trabajo presencial','')}</td>
                <td>{r.get('Descripción trabajo independiente','')}</td>
            </tr>
            """

    eval_rows = ""
    if evaluaciones_df is not None and not evaluaciones_df.empty:
        for _, r in evaluaciones_df.iterrows():
            eval_rows += f"""
            <tr>
                <td>{r.get('Tipo de evaluación','')}</td>
                <td>{r.get('Procedimiento de evaluación','')}</td>
                <td style="text-align:center;">{r.get('Valor (%)','')}%</td>
                <td style="text-align:center;">{r.get('Fecha de realización','')}</td>
                <td>{r.get('Unidad relacionada','')}</td>
            </tr>
            """

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>FD-GC71 Guía Didáctica - {datos.get('asignatura','')}</title>
    <style>
        @page {{ size: letter; margin: 1.8cm; }}
        body {{ font-family: Arial, sans-serif; color: #1a1a1a; font-size: 11pt; line-height: 1.4; margin: 0; padding: 20px; }}
        .header-table {{ width: 100%; border-collapse: collapse; margin-bottom: 20px; border: 2px solid #002060; }}
        .header-table td {{ border: 1px solid #002060; padding: 8px; text-align: center; }}
        .header-title {{ font-weight: bold; font-size: 14pt; color: #002060; }}
        h2 {{ color: #002060; font-size: 12pt; border-bottom: 2px solid #002060; padding-bottom: 4px; margin-top: 24px; }}
        .data-table {{ width: 100%; border-collapse: collapse; margin-bottom: 16px; font-size: 9.5pt; }}
        .data-table th {{ background-color: #002060; color: white; border: 1px solid #002060; padding: 6px; }}
        .data-table td {{ border: 1px solid #cccccc; padding: 6px; }}
        .print-btn {{ display: block; width: 200px; margin: 0 auto 20px auto; padding: 10px; background: #1f4fd8; color: white; text-align: center; border-radius: 8px; font-weight: bold; text-decoration: none; cursor: pointer; }}
        @media print {{ .print-btn {{ display: none; }} }}
    </style>
</head>
<body>
    <button class="print-btn" onclick="window.print()">🖨️ Imprimir / Guardar PDF</button>

    <table class="header-table">
        <tr>
            <td style="width:25%; font-weight:bold; color:#002060;">POLITÉCNICO COLOMBIANO<br>JAIME ISAZA CADAVID</td>
            <td class="header-title">GUÍA DIDÁCTICA Y CONCERTACIÓN DE EVALUACIÓN<br>(FD-GC71)</td>
            <td style="width:20%; font-size:9pt;">Código: FD-GC71<br>Versión: 04</td>
        </tr>
    </table>

    <h2>1. IDENTIFICACIÓN DE LA ASIGNATURA</h2>
    <table class="data-table">
        <tr>
            <td><strong>Programa:</strong> {datos.get('programa','')}</td>
            <td><strong>Asignatura:</strong> {datos.get('asignatura','')}</td>
            <td><strong>Código:</strong> {datos.get('codigo','')}</td>
        </tr>
        <tr>
            <td><strong>Profesor:</strong> {datos.get('profesor','')}</td>
            <td><strong>Correo:</strong> {datos.get('correo','')}</td>
            <td><strong>Grupo:</strong> {datos.get('grupo','')}</td>
        </tr>
        <tr>
            <td><strong>Créditos:</strong> {datos.get('creditos','')}</td>
            <td><strong>HTP:</strong> {datos.get('htp','')} | <strong>HTI:</strong> {datos.get('hti','')}</td>
            <td><strong>Periodo:</strong> {datos.get('periodo','')}</td>
        </tr>
    </table>

    <h2>2. TEXTOS ACADÉMICOS BASE</h2>
    <p><strong>Justificación:</strong> {datos.get('justificacion','')}</p>
    <p><strong>Competencias:</strong> {datos.get('competencias','')}</p>
    <p><strong>Objetivo General:</strong> {datos.get('objetivo_general','')}</p>

    <h2>3. PLAN DE SESIONES</h2>
    <table class="data-table">
        <thead>
            <tr>
                <th>Unidad</th>
                <th>Sesión</th>
                <th>Fecha</th>
                <th>Horario</th>
                <th>Contenido</th>
                <th>Trabajo Presencial</th>
                <th>Trabajo Independiente</th>
            </tr>
        </thead>
        <tbody>
            {sesiones_rows}
        </tbody>
    </table>

    <h2>4. CONCERTACIÓN DE EVALUACIÓN</h2>
    <table class="data-table">
        <thead>
            <tr>
                <th>Tipo de Evaluación</th>
                <th>Procedimiento</th>
                <th>Valor (%)</th>
                <th>Fecha</th>
                <th>Unidad Relacionada</th>
            </tr>
        </thead>
        <tbody>
            {eval_rows}
        </tbody>
    </table>
</body>
</html>
"""
