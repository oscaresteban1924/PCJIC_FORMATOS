from __future__ import annotations

from typing import Any, Dict, Optional, Tuple

import pandas as pd

from src.database.connection import db_execute, read_sql_df
from src.security import ahora_iso, hash_password, validar_password_seguro, verificar_password


def crear_usuario(usuario: str, nombre: str, email: str, rol: str, password: str, debe_cambiar: bool = True) -> int:
    usr = usuario.strip().lower()
    validar_password_seguro(password)
    existing = db_execute("SELECT id FROM usuarios WHERE usuario=?", (usr,), fetchone=True)
    if existing:
        raise ValueError(f"El usuario '{usr}' ya existe.")
    salt, p_hash = hash_password(password)
    row = db_execute(
        """
        INSERT INTO usuarios (usuario, nombre_completo, email, rol, salt, password_hash, activo, debe_cambiar_clave, creado_en)
        VALUES (?, ?, ?, ?, ?, ?, 1, ?, ?)
        """,
        (usr, nombre.strip(), email.strip(), rol, salt, p_hash, 1 if debe_cambiar else 0, ahora_iso()),
    )
    res = db_execute("SELECT id FROM usuarios WHERE usuario=?", (usr,), fetchone=True)
    return int(res["id"]) if res else 0


def actualizar_usuario(user_id: int, nombre: str, email: str, rol: str, activo: bool):
    db_execute(
        """
        UPDATE usuarios
        SET nombre_completo=?, email=?, rol=?, activo=?, actualizado_en=?
        WHERE id=?
        """,
        (nombre.strip(), email.strip(), rol, 1 if activo else 0, ahora_iso(), user_id),
    )


def cambiar_password(user_id: int, password: str, forzar_cambio: bool = False):
    validar_password_seguro(password)
    salt, p_hash = hash_password(password)
    db_execute(
        """
        UPDATE usuarios
        SET salt=?, password_hash=?, debe_cambiar_clave=?, actualizado_en=?
        WHERE id=?
        """,
        (salt, p_hash, 1 if forzar_cambio else 0, ahora_iso(), user_id),
    )


def listar_usuarios() -> pd.DataFrame:
    df = read_sql_df("SELECT id, usuario, nombre_completo, email, rol, activo, debe_cambiar_clave, creado_en, ultimo_login FROM usuarios ORDER BY id DESC")
    if not df.empty and "activo" in df.columns:
        df["activo"] = df["activo"].apply(lambda v: "Sí" if int(v or 0) == 1 else "No")
    return df


def autenticar_usuario(usuario: str, password: str) -> Optional[Dict[str, Any]]:
    usr = usuario.strip().lower()
    row = db_execute("SELECT * FROM usuarios WHERE usuario=?", (usr,), fetchone=True)
    if not row:
        return None
    user_dict = dict(row)
    if int(user_dict.get("activo", 0)) != 1:
        raise ValueError("El usuario está desactivado.")
    if not verificar_password(password, str(user_dict["salt"]), str(user_dict["password_hash"])):
        return None
    db_execute("UPDATE usuarios SET ultimo_login=? WHERE id=?", (ahora_iso(), user_dict["id"]))
    return user_dict
