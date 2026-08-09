from __future__ import annotations

from typing import Any, Dict

from src.config import (
    APP_VERSION,
    DEFAULT_MAX_EVIDENCE_MB,
    TEMPLATE_GC71,
    TEMPLATE_GC72,
    _safe_int_secret,
    get_app_env,
    usar_postgres,
)
from src.database.connection import db_execute


def health_status() -> Dict[str, Any]:
    """Diagnóstico de estado y salud del sistema."""
    status = {
        "version": APP_VERSION,
        "app_env": get_app_env(),
        "database": "PostgreSQL" if usar_postgres() else "SQLite local",
        "database_ok": False,
        "templates_ok": TEMPLATE_GC71.exists() and TEMPLATE_GC72.exists(),
        "storage": "DB + local cache" if usar_postgres() else "SQLite/local",
        "max_evidence_mb": _safe_int_secret("MAX_EVIDENCE_MB", DEFAULT_MAX_EVIDENCE_MB),
    }
    try:
        row = db_execute("SELECT COUNT(*) AS n FROM usuarios", fetchone=True)
        status["database_ok"] = row is not None
    except Exception as exc:
        status["database_error"] = str(exc)[:300]
    return status
