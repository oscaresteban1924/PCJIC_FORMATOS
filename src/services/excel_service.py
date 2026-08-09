from __future__ import annotations

import io
import re
from typing import Any, Dict, Iterable, List, Optional, Tuple

import pandas as pd
import xlsxwriter

from src.config import COLUMNAS_EVALUACIONES, COLUMNAS_GC72


def _normalizar_texto(texto: str) -> str:
    return re.sub(r"\s+", " ", str(texto or "").lower()).strip()


def porcentaje(numerador: Any, denominador: Any) -> str:
    try:
        num = float(numerador)
        den = float(denominador)
        if den == 0:
            return "0%"
        return f"{(num / den) * 100:.1f}%".replace(".0%", "%")
    except Exception:
        return "0%"


def normalizar_columnas_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def limpiar_numero(valor: Any) -> Optional[float]:
    if valor is None:
        return None
    if isinstance(valor, (int, float)):
        return float(valor) if not pd.isna(valor) else None
    txt = str(valor).strip().replace(",", ".")
    if not txt:
        return None
    try:
        return float(txt)
    except Exception:
        return None


def limpiar_df(df: pd.DataFrame, columnas: List[str]) -> pd.DataFrame:
    if df is None:
        return pd.DataFrame(columns=columnas)
    res = df.copy()
    for col in columnas:
        if col not in res.columns:
            res[col] = ""
    return res[columnas]


def _file_bytes(uploaded_file: Any) -> bytes:
    if uploaded_file is None:
        return b""
    try:
        return uploaded_file.getvalue()
    except Exception:
        pos = uploaded_file.tell()
        data = uploaded_file.read()
        try:
            uploaded_file.seek(pos)
        except Exception:
            pass
        return data


def _unique_columns(cols: Iterable[Any]) -> List[str]:
    seen: Dict[str, int] = {}
    out: List[str] = []
    for i, c in enumerate(cols):
        name = str(c).strip()
        if not name or name.lower() in {"nan", "none", "unnamed"}:
            name = f"Columna_{i+1}"
        base = name
        k = seen.get(base, 0)
        if k:
            name = f"{base}_{k+1}"
        seen[base] = k + 1
        out.append(name)
    return out


def _smart_header_row(raw: pd.DataFrame) -> int:
    if raw is None or raw.empty:
        return 0
    keywords = [
        "DOCUMENTO", "CEDULA", "CÉDULA", "IDENTIFIC", "CARN", "CODIGO", "CÓDIGO",
        "NOMBRE", "APELLIDO", "ESTUDIANTE", "CORREO", "EMAIL", "NOTA", "CALIFIC",
        "ESTADO", "OBSERV", "GRUPO", "PROGRAMA", "ASIGNATURA"
    ]
    best_idx, best_score = 0, -1
    max_rows = min(len(raw), 25)
    for idx in range(max_rows):
        vals = [_normalizar_texto(v) for v in raw.iloc[idx].tolist() if not pd.isna(v)]
        text = " | ".join(vals)
        non_empty = sum(1 for v in vals if str(v).strip())
        score = sum(1 for k in keywords if _normalizar_texto(k) in text) + min(non_empty, 8) * 0.08
        numeric_like = sum(1 for v in vals if re.fullmatch(r"[0-9., -]+", str(v).strip() or ""))
        score -= numeric_like * 0.05
        if score > best_score:
            best_idx, best_score = idx, score
    return int(best_idx)


def leer_tabla_excel(uploaded_file: Any) -> pd.DataFrame:
    if uploaded_file is None:
        return pd.DataFrame()
    data = _file_bytes(uploaded_file)
    if not data:
        return pd.DataFrame()
    nombre = getattr(uploaded_file, "name", "archivo").lower()
    if nombre.endswith(".csv"):
        return pd.read_csv(io.BytesIO(data))
    if nombre.endswith((".xlsx", ".xls")):
        return pd.read_excel(io.BytesIO(data), sheet_name=0, header=None)
    raise ValueError("Formato no soportado. Use .xls, .xlsx o .csv.")


def detectar_fila_encabezado(df: pd.DataFrame, palabras_clave: List[str]) -> int:
    claves = [_normalizar_texto(p) for p in palabras_clave]
    for idx, row in df.iterrows():
        texto = " | ".join(_normalizar_texto(v) for v in row.values if not pd.isna(v))
        coincidencias = sum(1 for clave in claves if clave in texto)
        if coincidencias >= max(1, min(2, len(claves))):
            return int(idx)
    return 0


def leer_listado_estudiantes(uploaded_file: Any) -> pd.DataFrame:
    if uploaded_file is None:
        return pd.DataFrame(columns=["Nombre completo", "Documento", "Correo", "Plan", "Observación", "Estado"])
    data = _file_bytes(uploaded_file)
    if not data:
        return pd.DataFrame(columns=["Nombre completo", "Documento", "Correo", "Plan", "Observación", "Estado"])
    nombre = getattr(uploaded_file, "name", "archivo").lower()
    if nombre.endswith(".csv"):
        raw = pd.read_csv(io.BytesIO(data), header=None)
    else:
        raw = pd.read_excel(io.BytesIO(data), sheet_name=0, header=None)
    header_row = detectar_fila_encabezado(raw, ["NOMBRE COMPLETO", "DOCUMENTO", "CORREO"])
    df = raw.iloc[header_row + 1:].copy()
    df.columns = [str(c).strip() for c in raw.iloc[header_row].tolist()]
    df = normalizar_columnas_df(df)
    rename = {}
    for col in df.columns:
        n = _normalizar_texto(col)
        if "NOMBRE" in n:
            rename[col] = "Nombre completo"
        elif "DOCUMENTO" in n or "CEDULA" in n or "CARN" in n:
            rename[col] = "Documento"
        elif "CORREO" in n or "EMAIL" in n:
            rename[col] = "Correo"
        elif n == "PLAN" or "PROGRAMA" in n:
            rename[col] = "Plan"
        elif "OBSERV" in n:
            rename[col] = "Observación"
    df = df.rename(columns=rename)
    for col in ["Nombre completo", "Documento", "Correo", "Plan", "Observación"]:
        if col not in df.columns:
            df[col] = ""
    df = df[["Nombre completo", "Documento", "Correo", "Plan", "Observación"]]
    df = df.dropna(how="all").copy()
    df = df[df["Nombre completo"].astype(str).str.strip().ne("")].reset_index(drop=True)
    df["Documento"] = df["Documento"].apply(lambda x: "" if pd.isna(x) else str(x).replace(".0", "").strip())
    df["Estado"] = df["Observación"].apply(lambda x: "Desertó" if re.search(r"DESERT|RETI|CANCEL", _normalizar_texto(x)) else "Activo")
    return df


def leer_excel_inteligente(uploaded_file: Any) -> Dict[str, pd.DataFrame]:
    if uploaded_file is None:
        return {}
    name = getattr(uploaded_file, "name", "archivo").lower()
    data = _file_bytes(uploaded_file)
    if not data:
        return {}
    bio = io.BytesIO(data)
    tablas: Dict[str, pd.DataFrame] = {}
    if name.endswith(".csv"):
        raw = pd.read_csv(bio, header=None, dtype=object)
        header = _smart_header_row(raw)
        df = raw.iloc[header + 1:].copy()
        df.columns = _unique_columns(raw.iloc[header].tolist())
        df = df.dropna(how="all").reset_index(drop=True)
        tablas["CSV"] = df
        return tablas
    xls = pd.ExcelFile(bio)
    for sheet in xls.sheet_names:
        raw = pd.read_excel(xls, sheet_name=sheet, header=None, dtype=object)
        if raw.empty:
            continue
        header = _smart_header_row(raw)
        df = raw.iloc[header + 1:].copy()
        df.columns = _unique_columns(raw.iloc[header].tolist())
        df = df.dropna(how="all").reset_index(drop=True)
        df = df.loc[:, [c for c in df.columns if not df[c].isna().all()]]
        if not df.empty:
            tablas[str(sheet)] = df
    return tablas


def leer_calificaciones(uploaded_file: Any) -> Tuple[pd.DataFrame, pd.DataFrame]:
    if uploaded_file is None:
        return pd.DataFrame(), pd.DataFrame()
    name = getattr(uploaded_file, "name", "archivo").lower()
    if not name.endswith((".xlsx", ".xls", ".csv")):
        raise ValueError("Formato de calificaciones no soportado.")
    data = _file_bytes(uploaded_file)
    if not data:
        return pd.DataFrame(), pd.DataFrame()
    if name.endswith(".csv"):
        return normalizar_columnas_df(pd.read_csv(io.BytesIO(data))), pd.DataFrame()
    xls = pd.ExcelFile(io.BytesIO(data))
    hoja_cal = None
    hoja_eval = None
    for sheet in xls.sheet_names:
        n = _normalizar_texto(sheet)
        if "CALIFIC" in n:
            hoja_cal = sheet
        if "EVALU" in n:
            hoja_eval = sheet
    if hoja_cal is None:
        hoja_cal = xls.sheet_names[0]
    cal = pd.read_excel(xls, sheet_name=hoja_cal)
    ev = pd.read_excel(xls, sheet_name=hoja_eval) if hoja_eval else pd.DataFrame()
    return normalizar_columnas_df(cal), normalizar_columnas_df(ev)


def encontrar_columna(df: pd.DataFrame, patrones: List[str]) -> Optional[str]:
    patrones_norm = [_normalizar_texto(p) for p in patrones]
    for col in df.columns:
        n = _normalizar_texto(col)
        if any(p in n for p in patrones_norm):
            return col
    return None


def contar_por_estado_y_nota(df: pd.DataFrame, columna_nota: Optional[str], corte_aprobacion: float, estado_col: Optional[str]) -> Tuple[int, int]:
    if df.empty or not columna_nota or columna_nota not in df.columns:
        return 0, 0
    notas = pd.to_numeric(df[columna_nota], errors="coerce")
    activo = pd.Series([True] * len(df), index=df.index)
    if estado_col and estado_col in df.columns:
        activo = ~df[estado_col].astype(str).apply(lambda x: bool(re.search(r"DESERT|RETI|CANCEL", _normalizar_texto(x))))
    validas = notas.notna() & activo
    aprueban = int((notas[validas] >= corte_aprobacion).sum())
    reprueban = int((notas[validas] < corte_aprobacion).sum())
    return aprueban, reprueban


def resumen_gc72_desde_archivos(
    listado_df: pd.DataFrame,
    calificaciones_df: pd.DataFrame,
    evaluaciones_df: pd.DataFrame,
    codigo: str,
    grupo: str,
    asignatura: str,
    avance_contenido: float,
    porcentaje_evaluado_manual: float,
    corte_aprobacion: float = 3.0,
) -> pd.DataFrame:
    matriculados = int(len(listado_df)) if listado_df is not None and not listado_df.empty else int(len(calificaciones_df)) if calificaciones_df is not None else 0
    estado_col_listado = "Estado" if listado_df is not None and "Estado" in listado_df.columns else None
    desertaron = 0
    if listado_df is not None and not listado_df.empty and estado_col_listado:
        desertaron = int(listado_df[estado_col_listado].astype(str).apply(lambda x: bool(re.search(r"DESERT|RETI|CANCEL", _normalizar_texto(x)))).sum())
    
    cal_df = calificaciones_df if calificaciones_df is not None else pd.DataFrame()
    estado_col = encontrar_columna(cal_df, ["Estado", "Observación", "Observacion"])
    parcial_col = encontrar_columna(cal_df, ["Nota parcial", "Parcial", "Corte 1", "Primer corte"])
    acumulada_col = encontrar_columna(cal_df, ["Nota acumulada", "Acumulada", "Definitiva", "Nota final", "Final"])
    if acumulada_col is None:
        acumulada_col = parcial_col

    apr_par, rep_par = contar_por_estado_y_nota(cal_df, parcial_col, corte_aprobacion, estado_col)
    apr_fecha, rep_fecha = contar_por_estado_y_nota(cal_df, acumulada_col, corte_aprobacion, estado_col)

    porcentaje_evaluado = porcentaje_evaluado_manual
    if evaluaciones_df is not None and not evaluaciones_df.empty:
        col_valor = encontrar_columna(evaluaciones_df, ["Valor", "%"])
        if col_valor:
            valores = pd.to_numeric(evaluaciones_df[col_valor], errors="coerce").fillna(0)
            porcentaje_evaluado = float(min(100, valores.sum())) if valores.sum() > 0 else porcentaje_evaluado_manual

    data = [{
        "Código": codigo,
        "Grupo": grupo,
        "Asignatura": asignatura,
        "% Avance en contenido": avance_contenido,
        "% Evaluado": porcentaje_evaluado,
        "Estudiantes matriculados": matriculados,
        "Desertaron N°": desertaron,
        "Desertaron %": porcentaje(desertaron, matriculados),
        "Aprueban evaluación parcial N°": apr_par,
        "Aprueban evaluación parcial %": porcentaje(apr_par, matriculados),
        "Reprueban evaluación parcial N°": rep_par,
        "Reprueban evaluación parcial %": porcentaje(rep_par, matriculados),
        "Aprueban a la fecha N°": apr_fecha,
        "Aprueban a la fecha %": porcentaje(apr_fecha, matriculados),
        "Reprueban a la fecha N°": rep_fecha,
        "Reprueban a la fecha %": porcentaje(rep_fecha, matriculados),
    }]
    return pd.DataFrame(data, columns=COLUMNAS_GC72)


ROLE_PATTERNS: Dict[str, List[str]] = {
    "documento": ["DOCUMENTO", "CEDULA", "CÉDULA", "IDENTIFIC", "ID", "CARN", "CODIGO", "CÓDIGO"],
    "nombre": ["NOMBRE", "ESTUDIANTE", "ALUMNO", "APELLIDO", "NOMBRES"],
    "correo": ["CORREO", "EMAIL", "MAIL", "E-MAIL"],
    "programa": ["PROGRAMA", "PLAN", "CARRERA"],
    "grupo": ["GRUPO", "CURSO"],
    "estado": ["ESTADO", "SITUACION", "SITUACIÓN", "OBSERV", "CONDIC"],
    "nota_parcial": ["PARCIAL", "CORTE", "SEGUIMIENTO", "NOTA PAR", "PROMEDIO PAR"],
    "nota_final": ["FINAL", "DEFINITIVA", "NOTA FINAL", "PROMEDIO FINAL", "CALIFICACION FINAL", "CALIFICACIÓN FINAL"],
    "nota": ["NOTA", "CALIFIC", "DEFINITIVA", "PROMEDIO", "TOTAL"],
}


def _guess_column(df: pd.DataFrame, role: str) -> Optional[str]:
    if df is None or df.empty:
        return None
    patterns = [_normalizar_texto(p) for p in ROLE_PATTERNS.get(role, [])]
    best_col, best_score = None, -1
    for col in df.columns:
        n = _normalizar_texto(col)
        score = 0
        for p in patterns:
            if p == n:
                score += 4
            elif p and p in n:
                score += 2
        sample = " ".join(_normalizar_texto(v) for v in df[col].head(20).tolist() if not pd.isna(v))
        if role == "correo" and "@" in sample:
            score += 3
        if role in {"nota", "nota_parcial", "nota_final"}:
            if re.search(r"DOCUMENTO|CEDULA|CÉDULA|IDENTIFIC|CARN|CODIGO|CÓDIGO|^ID$", n):
                score -= 4
            else:
                nums = pd.to_numeric(df[col].astype(str).str.replace(",", ".", regex=False), errors="coerce")
                ratio = float(nums.notna().mean()) if len(nums) else 0
                header_signal = any(p and p in n for p in patterns)
                if ratio > 0.45 and header_signal:
                    score += 2
                elif ratio > 0.60 and role == "nota":
                    score += 1
        if role == "estado" and re.search(r"ACTIVO|DESERT|RETI|CANCEL|MATRIC", sample):
            score += 2
        if score > best_score:
            best_col, best_score = col, score
    return best_col if best_score > 0 else None


def _mapping_automatico(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    return {role: _guess_column(df, role) for role in ROLE_PATTERNS.keys()}


def _clean_document_key(value: Any) -> str:
    if pd.isna(value):
        return ""
    s = str(value).strip().replace(".0", "")
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"[^A-Za-z0-9]", "", s)
    return s.upper()


def _clean_text(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _coerce_grade(value: Any) -> Optional[float]:
    if value is None or pd.isna(value):
        return None
    s = str(value).strip().replace("%", "").replace(",", ".")
    try:
        v = float(s)
    except Exception:
        return None
    if v > 100:
        return None
    if v > 5 and v <= 100:
        v = v / 20.0
    if v < 0 or v > 5:
        return None
    return round(v, 2)


def _estado_normalizado(value: Any) -> str:
    t = _normalizar_texto(value)
    if re.search(r"DESERT|RETI|CANCEL|ANUL|INACT|ABAND", t):
        return "Desertó"
    if re.search(r"APLAZ|SUSP", t):
        return "Suspendido"
    if re.search(r"MATRIC|ACTIV|INSCR|REGULAR|CURS", t):
        return "Activo"
    return _clean_text(value) or "Activo"


def normalizar_estudiantes_inteligente(df: pd.DataFrame, mapping: Dict[str, Optional[str]], nota_minima: float = 3.0) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=["Documento", "Documento llave", "Nombre completo", "Correo", "Programa", "Grupo", "Estado", "Nota parcial", "Nota final", "Resultado parcial", "Resultado final"])
    out = pd.DataFrame(index=df.index)

    def col(role: str) -> Optional[str]:
        c = mapping.get(role)
        return c if c and c in df.columns and c != "— No usar —" else None

    out["Documento"] = df[col("documento")].apply(_clean_text) if col("documento") else ""
    out["Documento llave"] = out["Documento"].apply(_clean_document_key)
    out["Nombre completo"] = df[col("nombre")].apply(_clean_text) if col("nombre") else ""
    out["Correo"] = df[col("correo")].apply(_clean_text) if col("correo") else ""
    out["Programa"] = df[col("programa")].apply(_clean_text) if col("programa") else ""
    out["Grupo"] = df[col("grupo")].apply(_clean_text) if col("grupo") else ""
    out["Estado"] = df[col("estado")].apply(_estado_normalizado) if col("estado") else "Activo"
    nota_parcial_col = col("nota_parcial") or col("nota")
    nota_final_col = col("nota_final") or col("nota")
    out["Nota parcial"] = df[nota_parcial_col].apply(_coerce_grade) if nota_parcial_col else None
    out["Nota final"] = df[nota_final_col].apply(_coerce_grade) if nota_final_col else None
    out["Resultado parcial"] = out["Nota parcial"].apply(lambda v: "Sin nota" if v is None or pd.isna(v) else ("Aprueba" if float(v) >= nota_minima else "Reprueba"))
    out["Resultado final"] = out["Nota final"].apply(lambda v: "Sin nota" if v is None or pd.isna(v) else ("Aprueba" if float(v) >= nota_minima else "Reprueba"))
    out["Calidad registro"] = "OK"
    out.loc[out["Documento llave"].eq(""), "Calidad registro"] = "Sin documento"
    out.loc[out["Nombre completo"].eq(""), "Calidad registro"] = out["Calidad registro"].where(out["Calidad registro"].ne("OK"), "Sin nombre")
    out = out.dropna(how="all").reset_index(drop=True)
    mask = out[["Documento llave", "Nombre completo", "Correo"]].astype(str).agg("".join, axis=1).str.strip().ne("")
    return out[mask].reset_index(drop=True)


def diagnostico_tabla_estudiantes(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    total = len(df) if df is not None else 0
    rows.append({"Tipo": "Info", "Hallazgo": "Registros procesados", "Cantidad": total, "Acción sugerida": "Validar muestra visual antes de descargar."})
    if df is None or df.empty:
        rows.append({"Tipo": "Error", "Hallazgo": "No hay registros útiles", "Cantidad": 0, "Acción sugerida": "Revise hoja, encabezado o archivo cargado."})
        return pd.DataFrame(rows)
    sin_doc = int(df["Documento llave"].eq("").sum()) if "Documento llave" in df.columns else 0
    if sin_doc:
        rows.append({"Tipo": "Alerta", "Hallazgo": "Estudiantes sin documento", "Cantidad": sin_doc, "Acción sugerida": "Completar documento o usar nombre como llave temporal."})
    sin_nombre = int(df["Nombre completo"].astype(str).str.strip().eq("").sum()) if "Nombre completo" in df.columns else 0
    if sin_nombre:
        rows.append({"Tipo": "Alerta", "Hallazgo": "Estudiantes sin nombre", "Cantidad": sin_nombre, "Acción sugerida": "Completar identificación nominal."})
    duplicados = int(df[df["Documento llave"].ne("")]["Documento llave"].duplicated().sum()) if "Documento llave" in df.columns else 0
    if duplicados:
        rows.append({"Tipo": "Error", "Hallazgo": "Documentos duplicados", "Cantidad": duplicados, "Acción sugerida": "Resolver duplicados antes de consolidar."})
    for c in ["Nota parcial", "Nota final"]:
        if c in df.columns:
            notas = pd.to_numeric(df[c], errors="coerce")
            fuera = int(((notas < 0) | (notas > 5)).sum())
            if fuera:
                rows.append({"Tipo": "Error", "Hallazgo": f"{c} fuera de escala 0-5", "Cantidad": fuera, "Acción sugerida": "Corregir escala o formato de notas."})
    desertores = int(df.get("Estado", pd.Series(dtype=str)).astype(str).str.contains("Desert", case=False, na=False).sum())
    if desertores:
        rows.append({"Tipo": "Info", "Hallazgo": "Registros marcados como desertores/retiros", "Cantidad": desertores, "Acción sugerida": "Validar que coincida con listado oficial."})
    return pd.DataFrame(rows)


def dataframe_to_xlsx_bytes(sheets: Dict[str, pd.DataFrame]) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        for name, df in sheets.items():
            sheet = re.sub(r"[\\/*?:\[\]]", "_", str(name))[:31] or "Hoja"
            (df if df is not None else pd.DataFrame()).to_excel(writer, sheet_name=sheet, index=False)
            ws = writer.sheets[sheet]
            wb = writer.book
            fmt_header = wb.add_format({"bold": True, "text_wrap": True, "valign": "top", "border": 1})
            for i, col in enumerate((df if df is not None else pd.DataFrame()).columns):
                ws.write(0, i, col, fmt_header)
                width = min(max(len(str(col)) + 4, 12), 42)
                ws.set_column(i, i, width)
    return bio.getvalue()


def excel_bytes_sheets(sheets: Dict[str, pd.DataFrame]) -> bytes:
    return dataframe_to_xlsx_bytes(sheets)


def crear_plantilla_evaluacion_xlsx(
    estudiantes_df: pd.DataFrame,
    evaluaciones_df: pd.DataFrame,
    datos: Dict[str, str],
    corte_aprobacion: float = 3.0,
) -> bytes:
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})
    fmt_title = workbook.add_format({"bold": True, "font_size": 14, "align": "center", "valign": "vcenter", "bg_color": "#D9EAF7", "border": 1})
    fmt_header = workbook.add_format({"bold": True, "bg_color": "#FFF2CC", "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True})
    fmt_cell = workbook.add_format({"border": 1, "valign": "vcenter"})
    fmt_num = workbook.add_format({"border": 1, "num_format": "0.00", "valign": "vcenter"})
    fmt_note = workbook.add_format({"border": 1, "font_color": "#666666", "italic": True, "text_wrap": True})

    ws = workbook.add_worksheet("Configuración")
    ws.merge_range("A1:D1", "Plantilla de evaluación generada desde FD-GC71", fmt_title)
    config_rows = [
        ("Programa", datos.get("programa", "")),
        ("Asignatura", datos.get("asignatura", "")),
        ("Código", datos.get("codigo", "")),
        ("Grupo", datos.get("grupo", "")),
        ("Periodo académico", datos.get("periodo", "")),
        ("Docente", datos.get("profesor", "")),
        ("Nota mínima aprobatoria", corte_aprobacion),
    ]
    ws.write_row("A3", ["Campo", "Valor"], fmt_header)
    for r, (k, v) in enumerate(config_rows, start=3):
        ws.write(r, 0, k, fmt_cell)
        ws.write(r, 1, v, fmt_num if isinstance(v, (int, float)) else fmt_cell)
    ws.set_column("A:A", 26)
    ws.set_column("B:D", 32)

    est = estudiantes_df.copy() if estudiantes_df is not None else pd.DataFrame()
    if est.empty:
        est = pd.DataFrame(columns=["Nombre completo", "Documento", "Correo", "Plan", "Observación", "Estado"])
    for col in ["Nombre completo", "Documento", "Correo", "Plan", "Observación", "Estado"]:
        if col not in est.columns:
            est[col] = ""
    est = est[["Nombre completo", "Documento", "Correo", "Plan", "Observación", "Estado"]]
    ws = workbook.add_worksheet("Estudiantes")
    ws.write_row(0, 0, est.columns.tolist(), fmt_header)
    for r, row in est.iterrows():
        for c, col in enumerate(est.columns):
            ws.write(r + 1, c, "" if pd.isna(row[col]) else row[col], fmt_cell)
    ws.autofilter(0, 0, max(1, len(est)), len(est.columns) - 1)
    ws.freeze_panes(1, 0)
    ws.set_column("A:A", 34)
    ws.set_column("B:B", 16)
    ws.set_column("C:C", 34)
    ws.set_column("D:D", 42)
    ws.set_column("E:F", 18)

    ev = limpiar_df(evaluaciones_df, COLUMNAS_EVALUACIONES)
    ws = workbook.add_worksheet("Evaluaciones")
    ws.write_row(0, 0, COLUMNAS_EVALUACIONES, fmt_header)
    total_valor = 0.0
    for r, row in ev.iterrows():
        for c, col in enumerate(COLUMNAS_EVALUACIONES):
            value = row.get(col, "")
            if col == "Fecha de realización" and hasattr(value, "strftime"):
                value = value.strftime("%d/%m/%Y")
            if col == "Valor (%)":
                val = limpiar_numero(value) or 0
                total_valor += val
                ws.write_number(r + 1, c, val, fmt_num)
            else:
                ws.write(r + 1, c, "" if pd.isna(value) else value, fmt_cell)
    ws.write(len(ev) + 2, 1, "Total concertado", fmt_header)
    ws.write_formula(len(ev) + 2, 2, f"=SUM(C2:C{len(ev)+1})", fmt_num)
    ws.write(len(ev) + 4, 0, "Nota", fmt_note)
    ws.merge_range(len(ev) + 4, 1, len(ev) + 4, 5, "El total de Valor (%) debe sumar 100. Las columnas de notas se generan automáticamente en la hoja Calificaciones.", fmt_note)
    ws.freeze_panes(1, 0)
    ws.set_column("A:A", 20)
    ws.set_column("B:B", 52)
    ws.set_column("C:C", 12)
    ws.set_column("D:F", 18)

    ws = workbook.add_worksheet("Calificaciones")
    fixed_headers = ["Nombre completo", "Documento", "Correo", "Estado"]
    eval_headers = [f"E{i+1} - {str(row.get('Tipo de evaluación', 'Evaluación')).strip()[:25]}" for i, (_, row) in enumerate(ev.iterrows())]
    headers = fixed_headers + eval_headers + ["Nota parcial", "Nota acumulada", "Aprueba a la fecha", "Observación docente"]
    ws.write_row(0, 0, headers, fmt_header)
    n = max(len(est), 1)
    for r in range(n):
        if r < len(est):
            ws.write(r + 1, 0, est.iloc[r].get("Nombre completo", ""), fmt_cell)
            ws.write(r + 1, 1, est.iloc[r].get("Documento", ""), fmt_cell)
            ws.write(r + 1, 2, est.iloc[r].get("Correo", ""), fmt_cell)
            ws.write(r + 1, 3, est.iloc[r].get("Estado", "Activo"), fmt_cell)
        else:
            for c in range(4):
                ws.write(r + 1, c, "", fmt_cell)
        for c in range(len(eval_headers)):
            ws.write_blank(r + 1, 4 + c, None, fmt_num)
        first_eval_col = 4
        last_eval_col = 4 + len(eval_headers) - 1
        excel_row = r + 2
        if eval_headers:
            weighted_terms = []
            for i in range(len(eval_headers)):
                col_letter = xlsxwriter.utility.xl_col_to_name(first_eval_col + i)
                peso_cell = f"Evaluaciones!$C${i+2}"
                weighted_terms.append(f"IF(ISNUMBER({col_letter}{excel_row}),{col_letter}{excel_row}*{peso_cell}/100,0)")
            formula = "=" + "+".join(weighted_terms)
            ws.write_formula(r + 1, last_eval_col + 1, formula, fmt_num)
            ws.write_formula(r + 1, last_eval_col + 2, formula, fmt_num)
            ws.write_formula(r + 1, last_eval_col + 3, f'=IF({xlsxwriter.utility.xl_col_to_name(last_eval_col + 2)}{excel_row}>=Configuración!$B$10,"Sí","No")', fmt_cell)
        else:
            ws.write_blank(r + 1, 4, None, fmt_num)
    ws.freeze_panes(1, 4)
    ws.autofilter(0, 0, n, len(headers) - 1)
    ws.set_column("A:A", 34)
    ws.set_column("B:B", 16)
    ws.set_column("C:C", 32)
    ws.set_column("D:D", 14)

    ws = workbook.add_worksheet("Resumen FD-GC72")
    resumen_headers = ["Indicador", "Valor", "Observación"]
    ws.write_row(0, 0, resumen_headers, fmt_header)
    cal_sheet = "Calificaciones"
    nota_parcial_col = xlsxwriter.utility.xl_col_to_name(4 + len(eval_headers)) if eval_headers else "E"
    nota_acum_col = xlsxwriter.utility.xl_col_to_name(5 + len(eval_headers)) if eval_headers else "F"
    estado_col = "D"
    filas = [
        ("Código", datos.get("codigo", ""), ""),
        ("Grupo", datos.get("grupo", ""), ""),
        ("Asignatura", datos.get("asignatura", ""), ""),
        ("Estudiantes matriculados", f"=COUNTA({cal_sheet}!A2:A{n+1})", "Desde listado de clase"),
        ("Desertaron N°", f'=COUNTIF({cal_sheet}!{estado_col}2:{estado_col}{n+1},"*Desert*")+COUNTIF({cal_sheet}!{estado_col}2:{estado_col}{n+1},"*Retir*")+COUNTIF({cal_sheet}!{estado_col}2:{estado_col}{n+1},"*Cancel*")', "Según columna Estado"),
        ("Aprueban evaluación parcial N°", f'=COUNTIF({cal_sheet}!{nota_parcial_col}2:{nota_parcial_col}{n+1},">="&Configuración!$B$10)', "Nota parcial >= mínima"),
        ("Reprueban evaluación parcial N°", f'=COUNTIF({cal_sheet}!{nota_parcial_col}2:{nota_parcial_col}{n+1},"<"&Configuración!$B$10)', "Nota parcial < mínima"),
        ("Aprueban a la fecha N°", f'=COUNTIF({cal_sheet}!{nota_acum_col}2:{nota_acum_col}{n+1},">="&Configuración!$B$10)', "Nota acumulada >= mínima"),
        ("Reprueban a la fecha N°", f'=COUNTIF({cal_sheet}!{nota_acum_col}2:{nota_acum_col}{n+1},"<"&Configuración!$B$10)', "Nota acumulada < mínima"),
        ("% Evaluado sugerido", "=Evaluaciones!C" + str(len(ev) + 3), "Puede ajustarse en la app"),
        ("% Avance en contenido", "", "Diligenciar en la app con base en sesiones realizadas / sesiones planificadas"),
    ]
    for r, (ind, val, obs) in enumerate(filas, start=1):
        ws.write(r, 0, ind, fmt_cell)
        if isinstance(val, str) and val.startswith("="):
            ws.write_formula(r, 1, val, fmt_num)
        else:
            ws.write(r, 1, val, fmt_cell)
        ws.write(r, 2, obs, fmt_cell)
    ws.set_column("A:A", 34)
    ws.set_column("B:B", 18)
    ws.set_column("C:C", 62)

    workbook.close()
    output.seek(0)
    return output.getvalue()


def crear_reporte_ejecutivo_xlsx(matriz_df: Optional[pd.DataFrame] = None) -> bytes:
    df = matriz_df if matriz_df is not None else pd.DataFrame()
    return dataframe_to_xlsx_bytes({"Cursos_Consolidado": df})
