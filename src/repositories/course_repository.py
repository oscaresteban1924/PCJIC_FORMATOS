from __future__ import annotations

import json
import hashlib
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

from src.config import COLUMNAS_EVALUACIONES, COLUMNAS_MODULOS, COLUMNAS_SESIONES, usar_postgres
from src.database.connection import conexion_db, db_execute, read_sql_df
from src.repositories.audit_repository import registrar_auditoria
from src.security import ahora_iso


def _dict_row(row: Any) -> Dict[str, Any]:
    if row is None:
        return {}
    if hasattr(row, "keys"):
        return {k: row[k] for k in row.keys()}
    if isinstance(row, dict):
        return row
    return dict(row)


def safe_json_loads(value: Any, default: Any = None) -> Any:
    if default is None:
        default = {}
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return default
    if isinstance(value, (dict, list)):
        return value
    try:
        return json.loads(value)
    except Exception:
        return default


def df_to_payload(df: pd.DataFrame) -> List[Dict[str, str]]:
    if df is None or df.empty:
        return []
    clean_df = df.fillna("")
    return clean_df.to_dict(orient="records")


def payload_to_df(payload: Any, columnas: Optional[List[str]] = None) -> pd.DataFrame:
    data = safe_json_loads(payload, []) if isinstance(payload, str) else (payload or [])
    try:
        df = pd.DataFrame(data)
    except Exception:
        df = pd.DataFrame(columns=columnas or [])
    if columnas:
        for c in columnas:
            if c not in df.columns:
                df[c] = ""
        df = df[columnas]
    return df


def hash_documental(value: Any) -> str:
    raw = json.dumps(value, ensure_ascii=False, sort_keys=True, default=str)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def cursos_visibles_query(user: Optional[Dict[str, Any]]) -> Tuple[str, Tuple[Any, ...]]:
    if not user:
        return "SELECT * FROM cursos ORDER BY id DESC", ()
    rol = user.get("rol", "")
    usr = user.get("usuario", "")
    if rol in ("Administrador", "Coordinador", "Consulta"):
        return "SELECT * FROM cursos ORDER BY id DESC", ()
    return "SELECT * FROM cursos WHERE propietario_usuario=? OR creado_por=? OR profesor LIKE ? ORDER BY id DESC", (usr, usr, f"%{usr}%")


def listar_cursos_visibles(user: Optional[Dict[str, Any]] = None) -> pd.DataFrame:
    sql, params = cursos_visibles_query(user)
    return read_sql_df(sql, params)


def get_curso(curso_id: int) -> Optional[Dict[str, Any]]:
    row = db_execute("SELECT * FROM cursos WHERE id=?", (int(curso_id),), fetchone=True)
    return _dict_row(row) if row else None


def esta_bloqueado(curso_id: int) -> Tuple[bool, Dict[str, Any]]:
    row = db_execute("SELECT * FROM curso_bloqueos WHERE curso_id=?", (int(curso_id),), fetchone=True)
    data = _dict_row(row)
    return bool(int(data.get("bloqueado", 0) or 0)), data


def registrar_workflow_evento(curso_id: int, evento: str, anterior: str = "", nuevo: str = "", resultado: str = "", detalle: str = "", user: Optional[Dict[str, Any]] = None):
    usr_dict = user or {}
    digest = ""
    db_execute(
        """
        INSERT INTO workflow_eventos(curso_id, evento, estado_anterior, estado_nuevo, resultado, detalle, hash_expediente, usuario, rol, creado_en)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (int(curso_id), evento, anterior, nuevo, resultado, detalle, digest, usr_dict.get("usuario", ""), usr_dict.get("rol", ""), ahora_iso()),
    )


def bloquear_curso(curso_id: int, motivo: str, user: Optional[Dict[str, Any]] = None):
    usr_dict = user or {}
    digest = hash_documental({"curso_id": curso_id, "motivo": motivo})
    db_execute("DELETE FROM curso_bloqueos WHERE curso_id=?", (int(curso_id),))
    db_execute(
        "INSERT INTO curso_bloqueos(curso_id, bloqueado, motivo, hash_bloqueo, bloqueado_por, bloqueado_en) VALUES (?, 1, ?, ?, ?, ?)",
        (int(curso_id), motivo, digest, usr_dict.get("usuario", ""), ahora_iso()),
    )
    registrar_workflow_evento(int(curso_id), "Bloqueo documental", detalle=motivo, resultado="Bloqueado", user=usr_dict)


def desbloquear_curso(curso_id: int, motivo: str, user: Optional[Dict[str, Any]] = None):
    db_execute("UPDATE curso_bloqueos SET bloqueado=0, motivo=?, bloqueado_en=? WHERE curso_id=?", (motivo, ahora_iso(), int(curso_id)))
    registrar_workflow_evento(int(curso_id), "Desbloqueo documental", detalle=motivo, resultado="Desbloqueado", user=user)


def upsert_curso(curso_id: Optional[int], datos: Dict[str, Any], payload: Optional[Dict[str, Any]] = None, user: Optional[Dict[str, Any]] = None) -> int:
    usr_dict = user or {}
    now = ahora_iso()
    if curso_id:
        bloqueado, info = esta_bloqueado(int(curso_id))
        if bloqueado and usr_dict.get("rol") not in ("Administrador",):
            raise RuntimeError(f"El expediente está bloqueado por aprobación/cierre. Motivo: {info.get('motivo','')}")

    payload_full = payload or {}
    payload_full["datos"] = datos
    payload_json = json.dumps(payload_full, ensure_ascii=False, default=str)

    params = (
        str(datos.get("codigo", "")),
        str(datos.get("grupo", "")),
        str(datos.get("asignatura", "Curso sin nombre")),
        str(datos.get("programa", "")),
        str(datos.get("periodo", "")),
        str(datos.get("profesor", "")),
        str(datos.get("correo", datos.get("email_profesor", ""))),
        str(datos.get("creditos", "")),
        float(datos.get("htp", 0) or 0),
        float(datos.get("hti", 0) or 0),
        str(datos.get("fecha_inicio", "")),
        str(datos.get("fecha_fin", "")),
        str(datos.get("estado", "Planeación")),
        float(datos.get("avance_contenido", 0) or 0),
        float(datos.get("avance_evaluado", 0) or 0),
        str(datos.get("propietario_usuario", usr_dict.get("usuario", ""))),
        payload_json,
    )

    conn = conexion_db()
    try:
        if curso_id:
            conn.execute(
                """
                UPDATE cursos
                SET codigo=?, grupo=?, asignatura=?, programa=?, periodo=?, profesor=?, email_profesor=?,
                    creditos=?, htp=?, hti=?, fecha_inicio=?, fecha_fin=?, estado=?, avance_contenido=?,
                    avance_evaluado=?, propietario_usuario=?, payload_json=?, actualizado_en=?
                WHERE id=?
                """,
                params + (now, int(curso_id)),
            )
            new_id = int(curso_id)
        else:
            if usar_postgres():
                cur = conn.execute(
                    """
                    INSERT INTO cursos (codigo, grupo, asignatura, programa, periodo, profesor, email_profesor, creditos, htp, hti, fecha_inicio, fecha_fin, estado, avance_contenido, avance_evaluado, propietario_usuario, payload_json, creado_por, creado_en, actualizado_en)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) RETURNING id
                    """,
                    params + (usr_dict.get("usuario", ""), now, now),
                )
                new_id = int(cur.fetchone()["id"])
            else:
                cur = conn.execute(
                    """
                    INSERT INTO cursos (codigo, grupo, asignatura, programa, periodo, profesor, email_profesor, creditos, htp, hti, fecha_inicio, fecha_fin, estado, avance_contenido, avance_evaluado, propietario_usuario, payload_json, creado_por, creado_en, actualizado_en)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    params + (usr_dict.get("usuario", ""), now, now),
                )
                new_id = int(cur.lastrowid)
        conn.commit()
        return new_id
    finally:
        conn.close()


def eliminar_curso(curso_id: int):
    db_execute("DELETE FROM cursos WHERE id=?", (int(curso_id),))


def siguiente_version(curso_id: int) -> int:
    row = db_execute("SELECT MAX(version_num) AS m FROM curso_versiones WHERE curso_id=?", (int(curso_id),), fetchone=True)
    val = dict(row).get("m") if row else None
    return (int(val) + 1) if val is not None else 1


def guardar_version_curso(curso_id: int, accion: str, nota: str = "", payload: Optional[Dict[str, Any]] = None, user: Optional[Dict[str, Any]] = None) -> int:
    usr_dict = user or {}
    v_num = siguiente_version(int(curso_id))
    if payload is None:
        c = get_curso(int(curso_id)) or {}
        payload = safe_json_loads(c.get("payload_json"), {})
    payload_json = json.dumps(payload, ensure_ascii=False, default=str)
    db_execute(
        """
        INSERT INTO curso_versiones(curso_id, version_num, accion, nota, creado_por, creado_en, payload_json)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (int(curso_id), v_num, accion, nota, usr_dict.get("usuario", ""), ahora_iso(), payload_json),
    )
    return v_num


def cambiar_estado_curso(curso_id: int, nuevo_estado: str, nota: str = "", user: Optional[Dict[str, Any]] = None):
    c = get_curso(int(curso_id))
    if not c:
        return
    anterior = str(c.get("estado", ""))
    db_execute("UPDATE cursos SET estado=?, actualizado_en=? WHERE id=?", (nuevo_estado, ahora_iso(), int(curso_id)))
    guardar_version_curso(int(curso_id), f"Cambio estado a {nuevo_estado}", nota, user=user)
    registrar_workflow_evento(int(curso_id), "Cambio de estado", anterior, nuevo_estado, "OK", nota, user=user)


def crear_observacion_curso(curso_id: int, prioridad: str, categoria: str, descripcion: str, user: Optional[Dict[str, Any]] = None) -> int:
    usr_dict = user or {}
    db_execute(
        """
        INSERT INTO curso_observaciones(curso_id, prioridad, categoria, descripcion, estado, creado_por, creado_en)
        VALUES (?, ?, ?, ?, 'Pendiente', ?, ?)
        """,
        (int(curso_id), prioridad, categoria, descripcion, usr_dict.get("usuario", ""), ahora_iso()),
    )
    return 1


def responder_observacion(obs_id: int, respuesta: str, user: Optional[Dict[str, Any]] = None):
    usr_dict = user or {}
    db_execute(
        """
        UPDATE curso_observaciones
        SET respuesta=?, estado='Atendida', respondido_por=?, respondido_en=?
        WHERE id=?
        """,
        (respuesta, usr_dict.get("usuario", ""), ahora_iso(), int(obs_id)),
    )


def observaciones_curso(curso_id: int) -> pd.DataFrame:
    return read_sql_df("SELECT * FROM curso_observaciones WHERE curso_id=? ORDER BY id DESC", (int(curso_id),))


def clonar_curso_a_nuevo_periodo(curso_id: int, nuevo_periodo: str, user: Optional[Dict[str, Any]] = None) -> int:
    """Clona un curso existente (syllabus, unidades, evaluaciones) a un nuevo periodo académico."""
    curso = get_curso(int(curso_id))
    if not curso:
        raise ValueError("Curso origen no encontrado.")

    payload = safe_json_loads(curso.get("payload_json"), {})
    datos = payload.get("datos", {}) or {}
    datos["periodo"] = nuevo_periodo

    datos_clone = {
        "codigo": curso.get("codigo", ""),
        "grupo": curso.get("grupo", ""),
        "asignatura": curso.get("asignatura", ""),
        "programa": curso.get("programa", ""),
        "periodo": nuevo_periodo,
        "profesor": curso.get("profesor", ""),
        "estado": "Planeación",
    }

    nuevo_id = upsert_curso(None, datos_clone, payload, user=user)
    registrar_auditoria("Clonar Curso", f"Clonado ID={curso_id} a nuevo ID={nuevo_id} (Periodo {nuevo_periodo})", user=user)
    return nuevo_id


def versiones_curso(curso_id: int) -> pd.DataFrame:
    return read_sql_df("SELECT id, version_num, accion, nota, creado_por, creado_en FROM curso_versiones WHERE curso_id=? ORDER BY version_num DESC", (int(curso_id),))
