from __future__ import annotations

import hashlib
import io
import json
import re
import zipfile
from datetime import date
from typing import Any, Dict, Optional, Tuple

import pandas as pd

from src.config import APP_VERSION
from src.repositories.course_repository import df_to_payload, get_curso, safe_json_loads, versiones_curso
from src.repositories.evidence_repository import evidencias_count
from src.security import ahora_iso
from src.services.calendar_service import build_ics_calendar
from src.services.docx_service import crear_gc71_docx
from src.services.excel_service import crear_plantilla_evaluacion_xlsx


def nombre_archivo_seguro(nombre: str, fecha: date | str, prefijo: str) -> str:
    base = re.sub(r"[^A-Za-zÁÉÍÓÚÜÑáéíóúüñ0-9]+", "_", str(nombre).strip()).strip("_") or "docente"
    f = fecha.strftime("%Y%m%d") if hasattr(fecha, "strftime") else re.sub(r"\D+", "", str(fecha))
    return f"{prefijo}_{base}_{f}"


def safe_filename(value: str, max_len: int = 80) -> str:
    name = re.sub(r"[^A-Za-z0-9_. -]", "_", str(value or "archivo")).strip(" ._")
    name = re.sub(r"_+", "_", name)
    return (name or "archivo")[:max_len]


def crear_paquete_curso_zip(
    datos: Dict[str, str],
    sesiones_df: pd.DataFrame,
    evaluaciones_df: pd.DataFrame,
    estudiantes_df: pd.DataFrame,
    representantes_df: Optional[pd.DataFrame] = None,
    curso_id: Optional[int] = None,
) -> bytes:
    reps = representantes_df if representantes_df is not None else pd.DataFrame()
    gc71 = crear_gc71_docx(datos, sesiones_df, evaluaciones_df, reps)
    excel = crear_plantilla_evaluacion_xlsx(estudiantes_df, evaluaciones_df, datos)
    ics = build_ics_calendar(sesiones_df, datos)
    ses_csv = (sesiones_df if sesiones_df is not None else pd.DataFrame()).to_csv(index=False).encode("utf-8-sig")
    ev_csv = (evaluaciones_df if evaluaciones_df is not None else pd.DataFrame()).to_csv(index=False).encode("utf-8-sig")

    payload = {
        "datos": datos,
        "sesiones": df_to_payload(sesiones_df),
        "evaluaciones": df_to_payload(evaluaciones_df),
        "estudiantes": df_to_payload(estudiantes_df),
        "representantes": df_to_payload(reps),
        "generado_en": ahora_iso(),
        "curso_id": curso_id,
    }
    buf = io.BytesIO()
    nombre_base = nombre_archivo_seguro(datos.get("asignatura", "curso"), date.today(), "FDGC_Paquete")
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(f"{nombre_base}/01_FD-GC71_Guia_Didactica.docx", gc71)
        z.writestr(f"{nombre_base}/02_Plantilla_Evaluacion.xlsx", excel)
        z.writestr(f"{nombre_base}/03_Calendario_Clases.ics", ics)
        z.writestr(f"{nombre_base}/04_Sesiones.csv", ses_csv)
        z.writestr(f"{nombre_base}/05_Evaluaciones.csv", ev_csv)
        z.writestr(f"{nombre_base}/06_Payload_Reutilizable.json", json.dumps(payload, ensure_ascii=False, indent=2, default=str).encode("utf-8"))
        z.writestr(f"{nombre_base}/LEAME.txt", "Paquete generado desde Gestor FD-GC71/FD-GC72 Enterprise.\n")
    return buf.getvalue()


def _hash_expediente(curso_id: int) -> Tuple[str, Dict[str, Any]]:
    curso = get_curso(int(curso_id)) or {}
    payload = safe_json_loads(curso.get("payload_json"), {})
    evid = evidencias_count(int(curso_id))
    versiones = versiones_curso(int(curso_id))
    data = {
        "app_version": APP_VERSION,
        "curso_id": int(curso_id),
        "codigo": curso.get("codigo", ""),
        "grupo": curso.get("grupo", ""),
        "asignatura": curso.get("asignatura", ""),
        "programa": curso.get("programa", ""),
        "periodo": curso.get("periodo", ""),
        "estado": curso.get("estado", ""),
        "actualizado_en": curso.get("actualizado_en", ""),
        "evidencias": evid,
        "versiones": len(versiones) if versiones is not None else 0,
        "payload_sha256": hashlib.sha256(json.dumps(payload, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest(),
    }
    digest = hashlib.sha256(json.dumps(data, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest()
    data["sha256_expediente"] = digest
    return digest, data
