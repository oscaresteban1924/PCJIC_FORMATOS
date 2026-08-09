from __future__ import annotations

import json
from typing import Any, Dict, Optional

import pandas as pd

from src.config import usar_postgres
from src.database.connection import conexion_db, db_execute, read_sql_df
from src.security import ahora_iso


def _dict_row(row: Any) -> Dict[str, Any]:
    if row is None:
        return {}
    if hasattr(row, "keys"):
        return {k: row[k] for k in row.keys()}
    if isinstance(row, dict):
        return row
    return dict(row)


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


def listar_asignaturas_base() -> pd.DataFrame:
    try:
        return read_sql_df("SELECT * FROM asignaturas_base WHERE activo=1 ORDER BY programa, nombre")
    except Exception:
        return pd.DataFrame()


def guardar_asignatura_base(data: Dict[str, Any], asignatura_id: Optional[int] = None, user: Optional[Dict[str, Any]] = None) -> int:
    usr_dict = user or {}
    now = ahora_iso()
    unidades_json = json.dumps(data.get("unidades", []), ensure_ascii=False, default=str)
    evaluaciones_json = json.dumps(data.get("evaluaciones", []), ensure_ascii=False, default=str)
    params = (
        data.get("codigo", ""),
        data.get("nombre", ""),
        data.get("programa", ""),
        data.get("area_formacion", ""),
        data.get("creditos", ""),
        limpiar_numero(data.get("htp", 0)) or 0,
        limpiar_numero(data.get("hti", 0)) or 0,
        data.get("tipo_asignatura", ""),
        data.get("justificacion", ""),
        data.get("competencias", ""),
        data.get("resultados", ""),
        data.get("objetivos", ""),
        data.get("metodologia", ""),
        data.get("ambientes", ""),
        data.get("medios", ""),
        data.get("bibliografia", ""),
        unidades_json,
        evaluaciones_json,
    )
    conn = conexion_db()
    try:
        if asignatura_id:
            conn.execute(
                """
                UPDATE asignaturas_base SET codigo=?, nombre=?, programa=?, area_formacion=?, creditos=?, htp=?, hti=?, tipo_asignatura=?, justificacion=?, competencias=?, resultados=?, objetivos=?, metodologia=?, ambientes=?, medios=?, bibliografia=?, unidades_json=?, evaluaciones_json=?, actualizado_en=? WHERE id=?
                """,
                params + (now, int(asignatura_id)),
            )
            new_id = int(asignatura_id)
        else:
            if usar_postgres():
                cur = conn.execute(
                    """
                    INSERT INTO asignaturas_base(codigo, nombre, programa, area_formacion, creditos, htp, hti, tipo_asignatura, justificacion, competencias, resultados, objetivos, metodologia, ambientes, medios, bibliografia, unidades_json, evaluaciones_json, creado_por, creado_en, actualizado_en)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) RETURNING id
                    """,
                    params + (usr_dict.get("usuario", ""), now, now),
                )
                new_id = int(cur.fetchone()["id"])
            else:
                cur = conn.execute(
                    """
                    INSERT INTO asignaturas_base(codigo, nombre, programa, area_formacion, creditos, htp, hti, tipo_asignatura, justificacion, competencias, resultados, objetivos, metodologia, ambientes, medios, bibliografia, unidades_json, evaluaciones_json, creado_por, creado_en, actualizado_en)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    params + (usr_dict.get("usuario", ""), now, now),
                )
                new_id = int(cur.lastrowid)
        conn.commit()
        return new_id
    finally:
        conn.close()


def get_asignatura_base(asignatura_id: int) -> Optional[Dict[str, Any]]:
    row = db_execute("SELECT * FROM asignaturas_base WHERE id=?", (int(asignatura_id),), fetchone=True)
    return _dict_row(row) if row else None
