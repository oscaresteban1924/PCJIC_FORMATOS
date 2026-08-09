from __future__ import annotations

from typing import Any, Dict, Optional

import pandas as pd

from src.database.connection import db_execute, read_sql_df
from src.security import ahora_iso


def registrar_artefacto(curso_id: Optional[int], tipo: str, nombre_archivo: str, mime: str = "", tamano: int = 0, b64_str: str = "", user: Optional[Dict[str, Any]] = None) -> int:
    usr_dict = user or {}
    db_execute(
        """
        INSERT INTO artefactos (curso_id, tipo, nombre_archivo, mime, tamano, creado_por, creado_en, contenido_b64)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (curso_id, tipo, nombre_archivo, mime, tamano, usr_dict.get("usuario", ""), ahora_iso(), b64_str),
    )
    return 1


def evidencias_count(curso_id: int) -> int:
    row = db_execute("SELECT COUNT(*) AS c FROM evidencias WHERE curso_id=?", (int(curso_id),), fetchone=True)
    return int(dict(row).get("c", 0)) if row else 0


def listar_evidencias(curso_id: int) -> pd.DataFrame:
    return read_sql_df("SELECT id, tipo, nombre_original, tamano, descripcion, subido_por, subido_en FROM evidencias WHERE curso_id=? ORDER BY id DESC", (int(curso_id),))
