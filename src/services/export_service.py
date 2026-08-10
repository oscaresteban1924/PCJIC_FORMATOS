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


def generar_codigo_qr_bytes(texto: str) -> bytes:
    """Genera una imagen PNG del código QR de verificación."""
    try:
        import qrcode
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_M,
            box_size=4,
            border=2,
        )
        qr.add_data(texto)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()
    except Exception:
        return b""


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

    # Hash de integridad
    raw_json = json.dumps(payload, ensure_ascii=False, sort_keys=True, default=str)
    digest = hashlib.sha256(raw_json.encode("utf-8")).hexdigest()
    qr_bytes = generar_codigo_qr_bytes(f"VERIFICACION-FDGC|ID={curso_id}|HASH={digest[:16]}")

    buf = io.BytesIO()
    nombre_base = nombre_archivo_seguro(datos.get("asignatura", "curso"), date.today(), "FDGC_Paquete")
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(f"{nombre_base}/01_FD-GC71_Guia_Didactica.docx", gc71)
        z.writestr(f"{nombre_base}/02_Plantilla_Evaluacion.xlsx", excel)
        z.writestr(f"{nombre_base}/03_Calendario_Clases.ics", ics)
        z.writestr(f"{nombre_base}/04_Sesiones.csv", ses_csv)
        z.writestr(f"{nombre_base}/05_Evaluaciones.csv", ev_csv)
        z.writestr(f"{nombre_base}/06_Payload_Reutilizable.json", raw_json.encode("utf-8"))
        if qr_bytes:
            z.writestr(f"{nombre_base}/07_Verificacion_QR.png", qr_bytes)
        z.writestr(f"{nombre_base}/LEAME.txt", f"Paquete generado desde Gestor FD-GC71/FD-GC72 Enterprise.\nHash SHA256: {digest}\n")
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


def generar_backup_sistema_json() -> bytes:
    """Genera un archivo JSON comprimido con la copia completa del sistema."""
    from src.database.connection import read_sql_df

    tables = ["usuarios", "cursos", "evidencias", "artefactos", "parametros_app", "asignaturas_base", "auditoria"]
    backup_data: Dict[str, Any] = {
        "metadatos": {
            "app_version": APP_VERSION,
            "fecha_creacion": ahora_iso(),
            "tablas_respaldadas": tables,
        },
        "datos": {},
    }

    for table in tables:
        try:
            df = read_sql_df(f"SELECT * FROM {table}")
            backup_data["datos"][table] = df.to_dict(orient="records")
        except Exception:
            backup_data["datos"][table] = []

    raw = json.dumps(backup_data, ensure_ascii=False, indent=2, default=str)
    return raw.encode("utf-8")


def restaurar_backup_sistema_json(json_bytes: bytes) -> Tuple[bool, str]:
    """Restaura el sistema a partir de una copia de seguridad JSON."""
    from src.database.connection import db_execute

    try:
        content = json.loads(json_bytes.decode("utf-8"))
        datos = content.get("datos", {})
        if not datos:
            return False, "El archivo de respaldo no contiene datos válidos."

        # Restaurar usuarios
        if "usuarios" in datos:
            for u in datos["usuarios"]:
                try:
                    db_execute(
                        "INSERT INTO usuarios (usuario, password_hash, salt, rol, nombre_completo, email, estado, debe_cambiar_clave) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                        (u.get("usuario"), u.get("password_hash"), u.get("salt"), u.get("rol"), u.get("nombre_completo"), u.get("email"), u.get("estado", "activo"), u.get("debe_cambiar_clave", 0)),
                    )
                except Exception:
                    db_execute(
                        "UPDATE usuarios SET rol=?, nombre_completo=?, email=? WHERE usuario=?",
                        (u.get("rol"), u.get("nombre_completo"), u.get("email"), u.get("usuario")),
                    )

        # Restaurar cursos
        if "cursos" in datos:
            for c in datos["cursos"]:
                try:
                    db_execute(
                        "INSERT INTO cursos (codigo, grupo, asignatura, programa, periodo, profesor, estado, payload_json, creado_en, actualizado_en) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                        (c.get("codigo"), c.get("grupo"), c.get("asignatura"), c.get("programa"), c.get("periodo"), c.get("profesor"), c.get("estado"), c.get("payload_json"), c.get("creado_en"), c.get("actualizado_en")),
                    )
                except Exception:
                    pass

        return True, "Copia de seguridad restaurada exitosamente."
    except Exception as e:
        return False, f"Error durante la restauración: {str(e)}"

