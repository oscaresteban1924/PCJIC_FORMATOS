from __future__ import annotations

from typing import Any, Dict, Optional

import pandas as pd

from src.database.connection import db_execute, read_sql_df
from src.security import ahora_iso


def registrar_auditoria(accion: str, detalle: str, user: Optional[Dict[str, Any]] = None):
    usr_dict = user or {}
    db_execute(
        """
        INSERT INTO auditoria (fecha, usuario, rol, accion, detalle)
        VALUES (?, ?, ?, ?, ?)
        """,
        (ahora_iso(), usr_dict.get("usuario", "sistema"), usr_dict.get("rol", "Sistema"), accion, detalle),
    )


def listar_auditoria(limit: int = 500) -> pd.DataFrame:
    return read_sql_df(f"SELECT fecha, usuario, rol, accion, detalle FROM auditoria ORDER BY id DESC LIMIT {int(limit)}")
