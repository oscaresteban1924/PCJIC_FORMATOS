from __future__ import annotations

import sqlite3
from typing import Any, List, Optional, Tuple

import pandas as pd

from src.config import (
    DATA_DIR,
    DB_PATH,
    _safe_int_secret,
    postgres_url_normalizada,
    usar_postgres,
)

try:
    import psycopg2
    from psycopg2.extras import RealDictCursor
except Exception:  # pragma: no cover
    psycopg2 = None
    RealDictCursor = None


def _traducir_sql(sql: str) -> str:
    """Traduce placeholders SQLite (?) a PostgreSQL (%s) para consultas parametrizadas simples."""
    return sql.replace("?", "%s") if usar_postgres() else sql


class PgConnectionAdapter:
    """Adaptador mínimo para conservar la API usada por la app con SQLite (Liskov Substitution Principle)."""

    def __init__(self):
        if psycopg2 is None:
            raise RuntimeError("psycopg2-binary no está instalado. Revise requirements.txt.")
        timeout = _safe_int_secret("DB_CONNECT_TIMEOUT", 8)
        self._conn = psycopg2.connect(postgres_url_normalizada(), connect_timeout=timeout)

    def execute(self, sql: str, params: Tuple[Any, ...] = tuple()):
        cur = self._conn.cursor(cursor_factory=RealDictCursor)
        cur.execute(_traducir_sql(sql), params or tuple())
        return cur

    def cursor(self):
        return self._conn.cursor(cursor_factory=RealDictCursor)

    def commit(self):
        self._conn.commit()

    def rollback(self):
        self._conn.rollback()

    def close(self):
        self._conn.close()


def conexion_db():
    """Conexión unificada. PostgreSQL si DATABASE_URL está configurado; SQLite en local/demo."""
    if usar_postgres():
        return PgConnectionAdapter()
    DATA_DIR.mkdir(exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def read_sql_df(sql: str, params: Tuple[Any, ...] = tuple()) -> pd.DataFrame:
    """Lectura robusta a DataFrame para SQLite/PostgreSQL."""
    conn = conexion_db()
    try:
        if usar_postgres():
            cur = conn.execute(sql, params)
            rows = cur.fetchall()
            return pd.DataFrame([dict(r) for r in rows])
        return pd.read_sql_query(sql, conn, params=params)
    finally:
        conn.close()


def db_execute(sql: str, params: Tuple[Any, ...] = tuple(), fetchone: bool = False, fetchall: bool = False):
    conn = conexion_db()
    try:
        cur = conn.execute(sql, params)
        result = None
        if fetchone:
            row = cur.fetchone()
            result = dict(row) if row is not None and usar_postgres() else row
        elif fetchall:
            rows = cur.fetchall()
            result = [dict(r) for r in rows] if usar_postgres() else rows
        conn.commit()
        return result
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
        raise
    finally:
        conn.close()
