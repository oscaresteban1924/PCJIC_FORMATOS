from __future__ import annotations

import re
from datetime import datetime, time
from typing import Any, Dict, Optional

import pandas as pd


def parse_time_value(value: Any) -> Optional[time]:
    if isinstance(value, time):
        return value
    if isinstance(value, datetime):
        return value.time()
    if value is None or pd.isna(value):
        return None
    texto = str(value).strip()
    if not texto:
        return None
    for fmt in ["%H:%M", "%H:%M:%S", "%I:%M %p", "%I:%M%p"]:
        try:
            return datetime.strptime(texto.upper(), fmt).time()
        except ValueError:
            pass
    return None


def build_ics_calendar(sesiones_df: pd.DataFrame, datos: Dict[str, str]) -> bytes:
    asignatura = datos.get("asignatura", "Clase")
    grupo = datos.get("grupo", "")
    lugar_default = datos.get("ambiente", "")
    stamp = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//Gestor FDGC//Calendario Academico//ES",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
    ]
    df = sesiones_df.reset_index(drop=True) if sesiones_df is not None else pd.DataFrame()
    for i, row in df.iterrows():
        fecha = pd.to_datetime(row.get("Fecha", ""), errors="coerce")
        if pd.isna(fecha):
            continue
        horario = str(row.get("Horario", "")).strip()
        m = re.search(r"(\d{1,2}:\d{2})\s*[-–]\s*(\d{1,2}:\d{2})", horario)
        inicio = "08:00" if not m else m.group(1)
        fin = "10:00" if not m else m.group(2)
        dtstart = datetime.combine(fecha.date(), parse_time_value(inicio) or time(8, 0)).strftime("%Y%m%dT%H%M%S")
        dtend = datetime.combine(fecha.date(), parse_time_value(fin) or time(10, 0)).strftime("%Y%m%dT%H%M%S")
        uid = f"fdgc-{datos.get('codigo','curso')}-{datos.get('grupo','grupo')}-{i+1}@local"
        summary = f"{asignatura} {grupo} - Sesión {row.get('N° sesión', i+1)}"
        desc = f"Unidad: {row.get('Unidad','')}\\nContenido: {row.get('Contenido por desarrollar','')}\\nTrabajo presencial: {row.get('Descripción del trabajo presencial','')}\\nTrabajo independiente: {row.get('Descripción trabajo independiente','')}"
        location = str(row.get("Lugar / ambiente", "") or lugar_default)
        lines += [
            "BEGIN:VEVENT",
            f"UID:{uid}",
            f"DTSTAMP:{stamp}",
            f"DTSTART:{dtstart}",
            f"DTEND:{dtend}",
            f"SUMMARY:{summary}",
            f"DESCRIPTION:{desc}",
            f"LOCATION:{location}",
            "END:VEVENT",
        ]
    lines.append("END:VCALENDAR")
    return "\r\n".join(lines).encode("utf-8")


def build_ics_reminders(evaluaciones_df: pd.DataFrame, datos: Dict[str, str]) -> bytes:
    """Genera un archivo iCalendar con alertas de vencimiento (VALARM) 2 días antes para cada evaluación."""
    asignatura = datos.get("asignatura", "Curso")
    grupo = datos.get("grupo", "")
    from datetime import timezone
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//Gestor FDGC//Recordatorios Evaluacion//ES",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
    ]
    df = evaluaciones_df.reset_index(drop=True) if evaluaciones_df is not None else pd.DataFrame()
    for i, row in df.iterrows():
        fecha = pd.to_datetime(row.get("Fecha de realización", ""), errors="coerce")
        if pd.isna(fecha):
            continue
        dtstart = datetime.combine(fecha.date(), time(8, 0)).strftime("%Y%m%dT%H%M%S")
        dtend = datetime.combine(fecha.date(), time(10, 0)).strftime("%Y%m%dT%H%M%S")
        uid = f"fdgc-eval-{datos.get('codigo','c')}-{i+1}@local"
        summary = f"⚠️ Recordatorio: {row.get('Tipo de evaluación','Evaluación')} - {asignatura} {grupo}"
        desc = f"Valor: {row.get('Valor (%)')}%\\nProcedimiento: {row.get('Procedimiento de evaluación','')}\\nCorte: {row.get('Corte','')}"

        lines += [
            "BEGIN:VEVENT",
            f"UID:{uid}",
            f"DTSTAMP:{stamp}",
            f"DTSTART:{dtstart}",
            f"DTEND:{dtend}",
            f"SUMMARY:{summary}",
            f"DESCRIPTION:{desc}",
            "BEGIN:VALARM",
            "TRIGGER:-P2D",
            "ACTION:DISPLAY",
            f"DESCRIPTION:Recordatorio 48h: {summary}",
            "END:VALARM",
            "END:VEVENT",
        ]
    lines.append("END:VCALENDAR")
    return "\r\n".join(lines).encode("utf-8")
