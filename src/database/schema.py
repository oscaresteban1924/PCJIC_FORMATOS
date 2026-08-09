from __future__ import annotations

import json
from datetime import datetime
from typing import Any, Dict

from src.config import (
    BACKUP_DIR,
    DATA_DIR,
    EVIDENCE_DIR,
    EXPORT_DIR,
    initial_admin_config,
    usar_postgres,
)
from src.database.connection import conexion_db, db_execute
from src.security import ahora_iso, hash_password


def add_column_if_missing(table: str, column: str, ddl_type: str):
    """Migración defensiva compatible con SQLite/PostgreSQL."""
    conn = conexion_db()
    try:
        if usar_postgres():
            cur = conn.execute(
                "SELECT column_name FROM information_schema.columns WHERE table_name=%s AND column_name=%s",
                (table, column),
            )
            exists = cur.fetchone() is not None
            if not exists:
                conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {ddl_type}")
        else:
            cur = conn.execute(f"PRAGMA table_info({table})")
            cols = [r[1] for r in cur.fetchall()]
            if column not in cols:
                conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {ddl_type}")
        conn.commit()
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
    finally:
        conn.close()


def seed_parametros_institucionales():
    """Valores iniciales institucionales."""
    defaults = {
        "INSTITUCION_NOMBRE": "Politécnico Colombiano Jaime Isaza Cadavid",
        "INSTITUCION_CODIGO": "PCJIC",
        "INSTITUCION_DEPENDENCIA": "Facultad de Ingeniería / Vicerrectoría Docente",
        "EVALUACION_SEGUIMIENTO_MAX": "40",
        "EVALUACION_PARCIAL_MAX": "40",
        "CORTE1_FECHA_MAX": "",
        "CORTE2_FECHA_MAX": "",
        "CORTE3_FECHA_MAX": "",
        "ALERTA_DESERCION_UMBRAL": "15",
        "EXIGIR_REPRESENTANTES_GC71": "true",
        "BLOQUEAR_APROBADOS": "true",
    }
    for k, v in defaults.items():
        db_execute(
            """
            INSERT INTO parametros_app(clave, valor, descripcion, actualizado_en)
            VALUES (?, ?, 'Parametro por defecto', ?)
            ON CONFLICT DO NOTHING
            """ if usar_postgres() else """
            INSERT OR IGNORE INTO parametros_app(clave, valor, descripcion, actualizado_en)
            VALUES (?, ?, 'Parametro por defecto', ?)
            """,
            (k, v, ahora_iso()),
        )


def init_db():
    """Inicializa la base de datos completa de la plataforma en modo local o PostgreSQL."""
    DATA_DIR.mkdir(exist_ok=True)
    EVIDENCE_DIR.mkdir(exist_ok=True)
    EXPORT_DIR.mkdir(exist_ok=True)
    BACKUP_DIR.mkdir(exist_ok=True)

    conn = conexion_db()
    try:
        if usar_postgres():
            conn.execute("""
                CREATE TABLE IF NOT EXISTS usuarios (
                    id SERIAL PRIMARY KEY,
                    usuario TEXT UNIQUE NOT NULL,
                    nombre_completo TEXT NOT NULL,
                    email TEXT,
                    rol TEXT NOT NULL,
                    salt TEXT NOT NULL,
                    password_hash TEXT NOT NULL,
                    activo INTEGER NOT NULL DEFAULT 1,
                    debe_cambiar_clave INTEGER NOT NULL DEFAULT 1,
                    creado_en TEXT NOT NULL,
                    actualizado_en TEXT,
                    ultimo_login TEXT
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS auditoria (
                    id SERIAL PRIMARY KEY,
                    fecha TEXT NOT NULL,
                    usuario TEXT,
                    rol TEXT,
                    accion TEXT NOT NULL,
                    detalle TEXT
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS cursos (
                    id SERIAL PRIMARY KEY,
                    codigo TEXT,
                    grupo TEXT,
                    asignatura TEXT NOT NULL,
                    programa TEXT,
                    periodo TEXT,
                    profesor TEXT,
                    email_profesor TEXT,
                    creditos TEXT,
                    htp DOUBLE PRECISION DEFAULT 0,
                    hti DOUBLE PRECISION DEFAULT 0,
                    fecha_inicio TEXT,
                    fecha_fin TEXT,
                    estado TEXT DEFAULT 'Planeación',
                    avance_contenido DOUBLE PRECISION DEFAULT 0,
                    avance_evaluado DOUBLE PRECISION DEFAULT 0,
                    propietario_usuario TEXT,
                    creado_por TEXT,
                    creado_en TEXT NOT NULL,
                    actualizado_en TEXT,
                    payload_json TEXT DEFAULT '{}'
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS evidencias (
                    id SERIAL PRIMARY KEY,
                    curso_id INTEGER REFERENCES cursos(id) ON DELETE SET NULL,
                    tipo TEXT,
                    nombre_original TEXT NOT NULL,
                    nombre_archivo TEXT NOT NULL,
                    mime TEXT,
                    tamano INTEGER DEFAULT 0,
                    descripcion TEXT,
                    subido_por TEXT,
                    subido_en TEXT NOT NULL,
                    contenido_b64 TEXT
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS artefactos (
                    id SERIAL PRIMARY KEY,
                    curso_id INTEGER REFERENCES cursos(id) ON DELETE SET NULL,
                    tipo TEXT NOT NULL,
                    nombre_archivo TEXT NOT NULL,
                    mime TEXT,
                    tamano INTEGER DEFAULT 0,
                    creado_por TEXT,
                    creado_en TEXT NOT NULL,
                    contenido_b64 TEXT
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS parametros_app (
                    clave TEXT PRIMARY KEY,
                    valor TEXT,
                    descripcion TEXT,
                    actualizado_en TEXT
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS curso_versiones (
                    id SERIAL PRIMARY KEY,
                    curso_id INTEGER REFERENCES cursos(id) ON DELETE CASCADE,
                    version_num INTEGER NOT NULL,
                    accion TEXT NOT NULL,
                    nota TEXT,
                    creado_por TEXT,
                    creado_en TEXT NOT NULL,
                    payload_json TEXT DEFAULT '{}'
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS curso_observaciones (
                    id SERIAL PRIMARY KEY,
                    curso_id INTEGER REFERENCES cursos(id) ON DELETE CASCADE,
                    prioridad TEXT DEFAULT 'Media',
                    categoria TEXT DEFAULT 'General',
                    descripcion TEXT NOT NULL,
                    respuesta TEXT,
                    estado TEXT DEFAULT 'Pendiente',
                    creado_por TEXT,
                    creado_en TEXT NOT NULL,
                    respondido_por TEXT,
                    respondido_en TEXT
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS asignaturas_base (
                    id SERIAL PRIMARY KEY,
                    codigo TEXT,
                    nombre TEXT NOT NULL,
                    programa TEXT,
                    area_formacion TEXT,
                    creditos TEXT,
                    htp DOUBLE PRECISION DEFAULT 0,
                    hti DOUBLE PRECISION DEFAULT 0,
                    tipo_asignatura TEXT,
                    justificacion TEXT,
                    competencias TEXT,
                    resultados TEXT,
                    objetivos TEXT,
                    metodologia TEXT,
                    ambientes TEXT,
                    medios TEXT,
                    bibliografia TEXT,
                    unidades_json TEXT DEFAULT '[]',
                    evaluaciones_json TEXT DEFAULT '[]',
                    activo INTEGER DEFAULT 1,
                    creado_por TEXT,
                    creado_en TEXT NOT NULL,
                    actualizado_en TEXT
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS workflow_eventos (
                    id SERIAL PRIMARY KEY,
                    curso_id INTEGER REFERENCES cursos(id) ON DELETE CASCADE,
                    evento TEXT NOT NULL,
                    estado_anterior TEXT,
                    estado_nuevo TEXT,
                    resultado TEXT,
                    detalle TEXT,
                    hash_expediente TEXT,
                    usuario TEXT,
                    rol TEXT,
                    creado_en TEXT NOT NULL
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS curso_bloqueos (
                    curso_id INTEGER PRIMARY KEY REFERENCES cursos(id) ON DELETE CASCADE,
                    bloqueado INTEGER DEFAULT 0,
                    motivo TEXT,
                    hash_bloqueo TEXT,
                    bloqueado_por TEXT,
                    bloqueado_en TEXT
                )
            """)
        else:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS usuarios (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    usuario TEXT UNIQUE NOT NULL,
                    nombre_completo TEXT NOT NULL,
                    email TEXT,
                    rol TEXT NOT NULL,
                    salt TEXT NOT NULL,
                    password_hash TEXT NOT NULL,
                    activo INTEGER NOT NULL DEFAULT 1,
                    debe_cambiar_clave INTEGER NOT NULL DEFAULT 1,
                    creado_en TEXT NOT NULL,
                    actualizado_en TEXT,
                    ultimo_login TEXT
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS auditoria (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    fecha TEXT NOT NULL,
                    usuario TEXT,
                    rol TEXT,
                    accion TEXT NOT NULL,
                    detalle TEXT
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS cursos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    codigo TEXT,
                    grupo TEXT,
                    asignatura TEXT NOT NULL,
                    programa TEXT,
                    periodo TEXT,
                    profesor TEXT,
                    email_profesor TEXT,
                    creditos TEXT,
                    htp REAL DEFAULT 0,
                    hti REAL DEFAULT 0,
                    fecha_inicio TEXT,
                    fecha_fin TEXT,
                    estado TEXT DEFAULT 'Planeación',
                    avance_contenido REAL DEFAULT 0,
                    avance_evaluado REAL DEFAULT 0,
                    propietario_usuario TEXT,
                    creado_por TEXT,
                    creado_en TEXT NOT NULL,
                    actualizado_en TEXT,
                    payload_json TEXT DEFAULT '{}'
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS evidencias (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    curso_id INTEGER,
                    tipo TEXT,
                    nombre_original TEXT NOT NULL,
                    nombre_archivo TEXT NOT NULL,
                    mime TEXT,
                    tamano INTEGER DEFAULT 0,
                    descripcion TEXT,
                    subido_por TEXT,
                    subido_en TEXT NOT NULL,
                    contenido_b64 TEXT,
                    FOREIGN KEY(curso_id) REFERENCES cursos(id)
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS artefactos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    curso_id INTEGER,
                    tipo TEXT NOT NULL,
                    nombre_archivo TEXT NOT NULL,
                    mime TEXT,
                    tamano INTEGER DEFAULT 0,
                    creado_por TEXT,
                    creado_en TEXT NOT NULL,
                    contenido_b64 TEXT,
                    FOREIGN KEY(curso_id) REFERENCES cursos(id)
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS parametros_app (
                    clave TEXT PRIMARY KEY,
                    valor TEXT,
                    descripcion TEXT,
                    actualizado_en TEXT
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS curso_versiones (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    curso_id INTEGER,
                    version_num INTEGER NOT NULL,
                    accion TEXT NOT NULL,
                    nota TEXT,
                    creado_por TEXT,
                    creado_en TEXT NOT NULL,
                    payload_json TEXT DEFAULT '{}',
                    FOREIGN KEY(curso_id) REFERENCES cursos(id)
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS curso_observaciones (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    curso_id INTEGER,
                    prioridad TEXT DEFAULT 'Media',
                    categoria TEXT DEFAULT 'General',
                    descripcion TEXT NOT NULL,
                    respuesta TEXT,
                    estado TEXT DEFAULT 'Pendiente',
                    creado_por TEXT,
                    creado_en TEXT NOT NULL,
                    respondido_por TEXT,
                    respondido_en TEXT,
                    FOREIGN KEY(curso_id) REFERENCES cursos(id)
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS asignaturas_base (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    codigo TEXT,
                    nombre TEXT NOT NULL,
                    programa TEXT,
                    area_formacion TEXT,
                    creditos TEXT,
                    htp REAL DEFAULT 0,
                    hti REAL DEFAULT 0,
                    tipo_asignatura TEXT,
                    justificacion TEXT,
                    competencias TEXT,
                    resultados TEXT,
                    objetivos TEXT,
                    metodologia TEXT,
                    ambientes TEXT,
                    medios TEXT,
                    bibliografia TEXT,
                    unidades_json TEXT DEFAULT '[]',
                    evaluaciones_json TEXT DEFAULT '[]',
                    activo INTEGER DEFAULT 1,
                    creado_por TEXT,
                    creado_en TEXT NOT NULL,
                    actualizado_en TEXT
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS workflow_eventos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    curso_id INTEGER,
                    evento TEXT NOT NULL,
                    estado_anterior TEXT,
                    estado_nuevo TEXT,
                    resultado TEXT,
                    detalle TEXT,
                    hash_expediente TEXT,
                    usuario TEXT,
                    rol TEXT,
                    creado_en TEXT NOT NULL,
                    FOREIGN KEY(curso_id) REFERENCES cursos(id)
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS curso_bloqueos (
                    curso_id INTEGER PRIMARY KEY,
                    bloqueado INTEGER DEFAULT 0,
                    motivo TEXT,
                    hash_bloqueo TEXT,
                    bloqueado_por TEXT,
                    bloqueado_en TEXT,
                    FOREIGN KEY(curso_id) REFERENCES cursos(id)
                )
            """)

        conn.commit()
    finally:
        conn.close()

    # Migraciones defensivas
    add_column_if_missing("evidencias", "contenido_b64", "TEXT")
    add_column_if_missing("artefactos", "contenido_b64", "TEXT")
    add_column_if_missing("cursos", "propietario_usuario", "TEXT")

    # Usuario admin inicial
    adm_user, adm_pass, adm_name, adm_email = initial_admin_config()
    existing = db_execute("SELECT id FROM usuarios WHERE usuario=?", (adm_user,), fetchone=True)
    if not existing:
        salt, p_hash = hash_password(adm_pass)
        db_execute(
            """
            INSERT INTO usuarios (usuario, nombre_completo, email, rol, salt, password_hash, activo, debe_cambiar_clave, creado_en)
            VALUES (?, ?, ?, 'Administrador', ?, ?, 1, 0, ?)
            """,
            (adm_user, adm_name, adm_email, salt, p_hash, ahora_iso()),
        )

    try:
        seed_parametros_institucionales()
    except Exception:
        pass
