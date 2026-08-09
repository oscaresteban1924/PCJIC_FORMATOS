from __future__ import annotations

import hashlib
import hmac
import re
import secrets
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

from src.config import (
    MODULO_APROBACION_BLOQUEANTE,
    MODULO_AUDITORIA_EXPEDIENTE,
    MODULO_BACKUP,
    MODULO_BANCO,
    MODULO_CENTRO,
    MODULO_COHERENCIA,
    MODULO_DIAGNOSTICO,
    MODULO_EVIDENCIAS,
    MODULO_EXPEDIENTE,
    MODULO_EXPORTACION_INSTITUCIONAL,
    MODULO_FLUJO,
    MODULO_INFORME_INSTITUCIONAL,
    MODULO_MOTOR,
    MODULO_PARAMETROS,
    MODULO_PLANEADOR,
    MODULO_REPORTES,
    MODULO_VALIDADOR,
    get_app_env,
)


def ahora_iso() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def hash_password(password: str, salt: Optional[str] = None) -> Tuple[str, str]:
    salt = salt or secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), 240_000)
    return salt, digest.hex()


def verificar_password(password: str, salt: str, password_hash: str) -> bool:
    _, digest = hash_password(password, salt)
    return hmac.compare_digest(digest, password_hash)


def validar_password_seguro(password: str):
    """Política de contraseña. En producción exige mayor rigor; en local conserva usabilidad."""
    min_len = 12 if get_app_env() == "production" else 8
    if len(password or "") < min_len:
        raise ValueError(f"La contraseña debe tener mínimo {min_len} caracteres.")
    if get_app_env() == "production":
        if not re.search(r"[A-Z]", password) or not re.search(r"[0-9]", password) or not re.search(r"[^a-zA-Z0-9]", password):
            raise ValueError("En producción la clave debe tener mayúscula, número y carácter especial.")


ROLES_PERMISOS: Dict[str, Dict[str, Any]] = {
    "Administrador": {
        "descripcion": "Control total: usuarios, perfiles, auditoría y generación de formatos.",
        "modulos": [
            "Inicio",
            MODULO_CENTRO,
            MODULO_EXPEDIENTE,
            MODULO_PLANEADOR,
            MODULO_FLUJO,
            MODULO_MOTOR,
            MODULO_REPORTES,
            MODULO_PARAMETROS,
            MODULO_BANCO,
            MODULO_COHERENCIA,
            MODULO_APROBACION_BLOQUEANTE,
            MODULO_AUDITORIA_EXPEDIENTE,
            MODULO_INFORME_INSTITUCIONAL,
            MODULO_EXPORTACION_INSTITUCIONAL,
            "FD-GC71 - Planeación",
            "FD-GC72 - Informe académico",
            MODULO_EVIDENCIAS,
            MODULO_VALIDADOR,
            MODULO_BACKUP,
            MODULO_DIAGNOSTICO,
            "Usuarios y perfiles",
            "Auditoría",
            "Mi cuenta",
            "Ayuda / flujo recomendado",
        ],
    },
    "Coordinador": {
        "descripcion": "Revisión académica: puede usar formatos, ver auditoría y acompañar cierres.",
        "modulos": [
            "Inicio",
            MODULO_CENTRO,
            MODULO_EXPEDIENTE,
            MODULO_PLANEADOR,
            MODULO_FLUJO,
            MODULO_MOTOR,
            MODULO_REPORTES,
            MODULO_BANCO,
            MODULO_COHERENCIA,
            MODULO_APROBACION_BLOQUEANTE,
            MODULO_AUDITORIA_EXPEDIENTE,
            MODULO_INFORME_INSTITUCIONAL,
            MODULO_EXPORTACION_INSTITUCIONAL,
            "FD-GC71 - Planeación",
            "FD-GC72 - Informe académico",
            MODULO_EVIDENCIAS,
            MODULO_VALIDADOR,
            MODULO_DIAGNOSTICO,
            "Auditoría",
            "Mi cuenta",
            "Ayuda / flujo recomendado",
        ],
    },
    "Docente": {
        "descripcion": "Operación docente: planeación, evaluación, informes y descarga de soportes.",
        "modulos": [
            "Inicio",
            MODULO_CENTRO,
            MODULO_EXPEDIENTE,
            MODULO_PLANEADOR,
            MODULO_FLUJO,
            MODULO_MOTOR,
            MODULO_COHERENCIA,
            MODULO_APROBACION_BLOQUEANTE,
            MODULO_AUDITORIA_EXPEDIENTE,
            "FD-GC71 - Planeación",
            "FD-GC72 - Informe académico",
            MODULO_EVIDENCIAS,
            MODULO_VALIDADOR,
            "Mi cuenta",
            "Ayuda / flujo recomendado",
        ],
    },
    "Consulta": {
        "descripcion": "Acceso de lectura al flujo y orientación operativa.",
        "modulos": ["Inicio", MODULO_CENTRO, MODULO_REPORTES, MODULO_INFORME_INSTITUCIONAL, "Mi cuenta", "Ayuda / flujo recomendado"],
    },
}


def _add_modulos_a_rol(rol: str, nuevos: List[str]):
    if rol not in ROLES_PERMISOS:
        return
    actuales = ROLES_PERMISOS[rol].setdefault("modulos", [])
    for m in nuevos:
        if m not in actuales:
            actuales.append(m)


def tiene_permiso(modulo: str, user: Optional[Dict[str, Any]] = None) -> bool:
    if user is None:
        try:
            import streamlit as st
            user = st.session_state.get("auth_user", {}) if hasattr(st, "session_state") else {}
        except Exception:
            user = {}
    if not user:
        return False
    rol = user.get("rol", "")
    return modulo in ROLES_PERMISOS.get(rol, {}).get("modulos", [])
