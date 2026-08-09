from __future__ import annotations

import math
import re
from datetime import date, datetime, time, timedelta
from typing import Any, Dict, Iterable, List, Optional, Tuple

import pandas as pd

from src.config import (
    COLUMNAS_EVALUACIONES,
    COLUMNAS_HORARIOS,
    COLUMNAS_MODULOS,
    COLUMNAS_SESIONES,
    DIAS_INV,
    DIAS_MAP,
)
from src.repositories.course_repository import (
    get_curso,
    listar_cursos_visibles,
    payload_to_df,
    safe_json_loads,
)
from src.repositories.evidence_repository import evidencias_count
from src.repositories.subject_repository import limpiar_numero
from src.services.calendar_service import parse_time_value


def limpiar_df(df: pd.DataFrame, columnas: List[str]) -> pd.DataFrame:
    if df is None:
        return pd.DataFrame(columns=columnas)
    res = df.copy()
    for col in columnas:
        if col not in res.columns:
            res[col] = ""
    return res[columnas]


def horas_entre(inicio: time, fin: time) -> float:
    di = datetime.combine(date.today(), inicio)
    df = datetime.combine(date.today(), fin)
    if df <= di:
        df += timedelta(days=1)
    return round((df - di).total_seconds() / 3600, 2)


def generar_fechas_clase(inicio: date, fin: date, horarios_df: pd.DataFrame, fechas_excluidas: Iterable[date] | None = None) -> pd.DataFrame:
    horarios = limpiar_df(horarios_df, COLUMNAS_HORARIOS)
    fechas_excluidas = set(fechas_excluidas or [])
    registros = []
    actual = inicio
    while actual <= fin:
        dia_nombre = DIAS_INV.get(actual.weekday(), "")
        for _, h in horarios.iterrows():
            dia = str(h.get("Día", "")).strip()
            if DIAS_MAP.get(dia, -1) == actual.weekday() and actual not in fechas_excluidas:
                hi = parse_time_value(h.get("Hora inicio")) or time(0, 0)
                hf = parse_time_value(h.get("Hora fin")) or time(0, 0)
                registros.append({
                    "Fecha": actual,
                    "Día": dia_nombre,
                    "Hora inicio": hi.strftime("%H:%M") if hi else "",
                    "Hora fin": hf.strftime("%H:%M") if hf else "",
                    "Horas": horas_entre(hi, hf) if hi and hf else 0,
                    "Lugar / ambiente": h.get("Lugar / ambiente", ""),
                })
        actual += timedelta(days=1)
    df = pd.DataFrame(registros)
    if df.empty:
        return pd.DataFrame(columns=["Fecha", "Día", "Hora inicio", "Hora fin", "Horas", "Lugar / ambiente"])
    df["orden"] = df["Fecha"].astype(str) + " " + df["Hora inicio"].astype(str)
    df = df.sort_values("orden").drop(columns=["orden"]).reset_index(drop=True)
    return df


def expandir_plan_sesiones(modulos_df: pd.DataFrame, fechas_clase_df: pd.DataFrame, criterio: str = "Horas presenciales") -> pd.DataFrame:
    modulos = limpiar_df(modulos_df, COLUMNAS_MODULOS)
    registros = []
    idx_fecha = 0
    n_sesion = 1
    for _, mod in modulos.iterrows():
        unidad = str(mod.get("Unidad", "")).strip() or f"Unidad {len(registros)+1}"
        contenido = str(mod.get("Contenido / tema central", "")).strip()
        trabajo_p = str(mod.get("Trabajo presencial", "")).strip()
        trabajo_i = str(mod.get("Trabajo independiente", "")).strip()
        horas_objetivo = limpiar_numero(mod.get("Horas presenciales")) or 0
        sesiones_objetivo = limpiar_numero(mod.get("Sesiones"))
        if criterio == "Sesiones" and sesiones_objetivo:
            n_bloques = int(max(1, round(sesiones_objetivo)))
        else:
            if fechas_clase_df.empty:
                n_bloques = int(max(1, math.ceil(horas_objetivo / 2))) if horas_objetivo else 1
            else:
                acumuladas = 0.0
                n_bloques = 0
                tmp_idx = idx_fecha
                while acumuladas < max(0.01, horas_objetivo) and tmp_idx < len(fechas_clase_df):
                    acumuladas += limpiar_numero(fechas_clase_df.iloc[tmp_idx].get("Horas")) or 0
                    n_bloques += 1
                    tmp_idx += 1
                n_bloques = max(1, n_bloques)
        for parte in range(1, n_bloques + 1):
            if idx_fecha < len(fechas_clase_df):
                f = fechas_clase_df.iloc[idx_fecha]
                fecha = f.get("Fecha")
                horario = f"{f.get('Hora inicio', '')} - {f.get('Hora fin', '')}".strip(" -")
                lugar = f.get("Lugar / ambiente", "")
            else:
                fecha = "Por programar"
                horario = "Por programar"
                lugar = ""
            sufijo = f" (parte {parte} de {n_bloques})" if n_bloques > 1 else ""
            registros.append({
                "Unidad": unidad,
                "N° sesión": n_sesion,
                "Fecha": fecha,
                "Horario": horario,
                "Contenido por desarrollar": f"{contenido}{sufijo}",
                "Descripción del trabajo presencial": trabajo_p,
                "Descripción trabajo independiente": trabajo_i,
                "Lugar / ambiente": lugar,
            })
            idx_fecha += 1
            n_sesion += 1
    return pd.DataFrame(registros, columns=COLUMNAS_SESIONES)


def validar_plan(
    modulos_df: pd.DataFrame,
    horarios_df: pd.DataFrame,
    sesiones_df: pd.DataFrame,
    evaluaciones_df: pd.DataFrame,
    datos: Dict[str, str],
) -> Tuple[List[str], List[str]]:
    alertas: List[str] = []
    recomendaciones: List[str] = []

    if modulos_df.empty:
        alertas.append("No se han definido unidades ni contenidos en la estructura del curso.")
    if horarios_df.empty:
        alertas.append("Falta configurar la distribución de días y horarios de clase.")
    if evaluaciones_df.empty:
        alertas.append("No se han configurado actividades evaluativas.")
    else:
        total_eval = sum(limpiar_numero(r.get("Valor (%)")) or 0 for _, r in evaluaciones_df.iterrows())
        if abs(total_eval - 100.0) > 0.1:
            alertas.append(f"La suma de porcentajes de evaluación es {total_eval:.1f}% (debe ser 100%).")

    htp = limpiar_numero(datos.get("htp")) or 0
    hti = limpiar_numero(datos.get("hti")) or 0
    creditos = limpiar_numero(datos.get("creditos")) or 0
    if creditos > 0 and (htp + hti) == 0:
        recomendaciones.append("Indique HTP y HTI para verificar la relación horaria por créditos.")

    return alertas, recomendaciones


def score_calidad_expediente(curso_id: int) -> Tuple[int, List[Tuple[str, str, int]]]:
    c = get_curso(int(curso_id)) or {}
    payload = safe_json_loads(c.get("payload_json"), {})
    datos = payload.get("datos", {}) or {}
    sesiones = payload.get("sesiones", [])
    evaluaciones = payload.get("evaluaciones", [])
    evid = evidencias_count(int(curso_id))

    criterios: List[Tuple[str, str, int]] = []
    score = 0

    if datos.get("asignatura") and datos.get("codigo") and datos.get("grupo"):
        criterios.append(("Identificación", "OK", 15))
        score += 15
    else:
        criterios.append(("Identificación", "Incompleta", 5))
        score += 5

    if sesiones:
        criterios.append(("Sesiones", "OK", 25))
        score += 25
    else:
        criterios.append(("Sesiones", "Faltan", 0))

    if evaluaciones:
        total = sum(limpiar_numero(e.get("Valor (%)")) or 0 for e in evaluaciones if isinstance(e, dict))
        if abs(total - 100.0) <= 0.1:
            criterios.append(("Evaluación", "Concertada 100%", 25))
            score += 25
        else:
            criterios.append(("Evaluación", f"Suma {total:.0f}%", 10))
            score += 10
    else:
        criterios.append(("Evaluación", "Falta", 0))

    if evid >= 2:
        criterios.append(("Soportes", "Completos", 20))
        score += 20
    elif evid == 1:
        criterios.append(("Soportes", "Parciales", 10))
        score += 10
    else:
        criterios.append(("Soportes", "Sin evidencias", 0))

    st_name = str(c.get("estado", "")).strip()
    if st_name in ("Aprobado", "Cerrado"):
        criterios.append(("Estado", "Cierre formal", 15))
        score += 15
    elif st_name in ("Reportado", "En revisión"):
        criterios.append(("Estado", "Avance positivo", 10))
        score += 10
    else:
        criterios.append(("Estado", "Inicial", 5))
        score += 5

    return score, criterios


def analizar_coherencia_curso(curso_id: int) -> Dict[str, Any]:
    c = get_curso(int(curso_id)) or {}
    payload = safe_json_loads(c.get("payload_json"), {})
    datos = payload.get("datos", {}) or {}
    sesiones = payload_to_df(payload.get("sesiones"), COLUMNAS_SESIONES)
    evaluaciones = payload_to_df(payload.get("evaluaciones"), COLUMNAS_EVALUACIONES)
    modulos = payload_to_df(payload.get("modulos"), COLUMNAS_MODULOS)

    hallazgos: List[Dict[str, Any]] = []
    checks: List[Tuple[str, int, int]] = []

    def add_check(comp: str, p: int, maxp: int, sev: str, hall: str, rec: str):
        checks.append((comp, max(0, min(p, maxp)), maxp))
        if p < maxp:
            hallazgos.append({"Componente": comp, "Severidad": sev, "Hallazgo": hall, "Recomendación": rec, "Puntos": p, "Máximo": maxp})

    # 1. Creditos vs HTP/HTI
    cred = limpiar_numero(datos.get("creditos", 0)) or 0
    htp = limpiar_numero(datos.get("htp", 0)) or 0
    hti = limpiar_numero(datos.get("hti", 0)) or 0
    if cred > 0:
        esperado_total = cred * 48
        actual_semanal = htp + hti
        if actual_semanal > 0:
            add_check("Horarios/Créditos", 20, 20, "Baja", "Relación horas/créditos correcta", "Mantener")
        else:
            add_check("Horarios/Créditos", 5, 20, "Alta", "Falta definir HTP/HTI", "Asignar HTP y HTI acordes al plan.")
    else:
        add_check("Horarios/Créditos", 10, 20, "Media", "Créditos no especificados", "Revisar créditos de la asignatura.")

    # 2. Evaluacion
    if not evaluaciones.empty:
        total_eval = sum(limpiar_numero(r.get("Valor (%)")) or 0 for _, r in evaluaciones.iterrows())
        if abs(total_eval - 100.0) <= 0.1:
            add_check("Evaluación", 25, 25, "Baja", "Concertación al 100%", "Mantener")
        else:
            add_check("Evaluación", 10, 25, "Alta", f"Suma {total_eval:.1f}% en lugar de 100%", "Ajustar la ponderación de las evaluaciones.")
    else:
        add_check("Evaluación", 0, 25, "Alta", "Sin actividades evaluativas", "Ingresar las evaluaciones concertadas.")

    # 3. Sesiones
    if not sesiones.empty:
        add_check("Plan de Sesiones", 25, 25, "Baja", f"{len(sesiones)} sesiones programadas", "OK")
    else:
        add_check("Plan de Sesiones", 5, 25, "Alta", "Sin sesiones en el plan", "Expandir y guardar plan de clases.")

    # 4. Soportes
    evid = evidencias_count(int(curso_id))
    if evid >= 1:
        add_check("Evidencias", 30, 30, "Baja", f"{evid} evidencias cargadas", "OK")
    else:
        add_check("Evidencias", 10, 30, "Media", "Sin evidencias cargadas", "Subir actas, listas o guías firmadas.")

    puntos_total = sum(c[1] for c in checks)
    max_total = sum(c[2] for c in checks)
    porcentaje = round((puntos_total / max_total) * 100, 1) if max_total else 0.0

    return {
        "curso_id": curso_id,
        "score": porcentaje,
        "puntos": puntos_total,
        "maximo": max_total,
        "checks": checks,
        "hallazgos": pd.DataFrame(hallazgos),
    }


def matriz_riesgo_cursos() -> pd.DataFrame:
    df_cursos = listar_cursos_visibles()
    if df_cursos.empty:
        return pd.DataFrame()
    rows = []
    for _, row in df_cursos.iterrows():
        cid = int(row["id"])
        score, _ = score_calidad_expediente(cid)
        evid = evidencias_count(cid)
        st_name = str(row.get("estado", ""))
        riesgo = "Bajo" if score >= 80 else ("Medio" if score >= 50 else "Alto")
        rows.append({
            "ID": cid,
            "Código": row.get("codigo", ""),
            "Grupo": row.get("grupo", ""),
            "Asignatura": row.get("asignatura", ""),
            "Profesor": row.get("profesor", ""),
            "Estado": st_name,
            "Score Calidad": score,
            "Evidencias": evid,
            "Nivel Riesgo": riesgo,
        })
    return pd.DataFrame(rows)


def generar_sugerencias_academicas(curso_id: int) -> Dict[str, str]:
    analisis = analizar_coherencia_curso(int(curso_id))
    score = analisis.get("score", 0)
    hallazgos_df = analisis.get("hallazgos", pd.DataFrame())

    sug = []
    if not hallazgos_df.empty:
        for _, row in hallazgos_df.iterrows():
            sug.append(f"- [{row.get('Componente')}] {row.get('Recomendación')}")
    else:
        sug.append("Expediente en perfecto estado metodológico y normativo.")

    return {
        "diagnostico": f"El curso presenta una calidad documental del {score}%.",
        "sugerencias": "\n".join(sug),
    }
