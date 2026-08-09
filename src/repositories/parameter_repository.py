from __future__ import annotations

from typing import Any, Dict, Optional

from src.database.connection import db_execute
from src.security import ahora_iso


def obtener_parametro(clave: str, default: str = "") -> str:
    row = db_execute("SELECT valor FROM parametros_app WHERE clave=?", (clave,), fetchone=True)
    if not row:
        return default
    val = dict(row).get("valor")
    return str(val) if val is not None else default


def actualizar_parametro(clave: str, valor: str, descripcion: str = ""):
    db_execute(
        """
        INSERT INTO parametros_app(clave, valor, descripcion, actualizado_en)
        VALUES (?, ?, ?, ?)
        ON CONFLICT(clave) DO UPDATE SET valor=EXCLUDED.valor, descripcion=EXCLUDED.descripcion, actualizado_en=EXCLUDED.actualizado_en
        """,
        (clave, valor, descripcion, ahora_iso()),
    )
