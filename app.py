from __future__ import annotations

"""Punto de entrada principal del Gestor Académico FD-GC71 / FD-GC72.

Refactorizado bajo arquitectura SOLID para mantenibilidad, escalabilidad y despliegue productivo.
Re-exporta símbolos públicos clave para mantener retrocompatibilidad total al 100% con scripts de prueba
(como scripts/smoke_test.py) y entornos de producción.
"""

# Configuración y Constantes
from src.config import (
    APP_DIR,
    APP_VERSION,
    COLUMNAS_EVALUACIONES,
    COLUMNAS_GC72,
    COLUMNAS_HORARIOS,
    COLUMNAS_MODULOS,
    COLUMNAS_SESIONES,
    DATA_DIR,
    DB_PATH,
    DEFAULT_MAX_EVIDENCE_MB,
    DIAS_INV,
    DIAS_MAP,
    EVIDENCE_DIR,
    EXPORT_DIR,
    LOGO_ICONTEC,
    LOGO_POLI,
    MAPA_TABLA_WORD_GC72,
    NUMERICAS_ENTERAS_GC72,
    NUMERICAS_PORCENTAJE_GC72,
    PREFORMAS_GC72,
    TEMPLATE_DIR,
    TEMPLATE_GC71,
    TEMPLATE_GC72,
    TEXTOS_PREDEFINIDOS_GC71,
    _database_url_es_placeholder,
    _get_secret_value,
    _safe_int_secret,
    get_app_env,
    get_database_url,
    initial_admin_config,
    max_evidence_bytes,
    postgres_url_normalizada,
    usar_postgres,
)

# Seguridad y Autenticación
from src.security import (
    ROLES_PERMISOS,
    ahora_iso,
    hash_password,
    tiene_permiso,
    validar_password_seguro,
    verificar_password,
)

# Base de datos y Esquema
from src.database.connection import (
    PgConnectionAdapter,
    _traducir_sql,
    conexion_db,
    db_execute,
    read_sql_df,
)
from src.database.schema import add_column_if_missing, init_db, seed_parametros_institucionales

# Repositorios de datos
from src.repositories.audit_repository import listar_auditoria, registrar_auditoria
from src.repositories.course_repository import (
    bloquear_curso,
    cambiar_estado_curso,
    crear_observacion_curso,
    cursos_visibles_query,
    desbloquear_curso,
    df_to_payload,
    eliminar_curso,
    esta_bloqueado,
    get_curso,
    guardar_version_curso,
    hash_documental,
    listar_cursos_visibles,
    observaciones_curso,
    payload_to_df,
    registrar_workflow_evento,
    responder_observacion,
    safe_json_loads,
    siguiente_version,
    upsert_curso,
    versiones_curso,
)
from src.repositories.evidence_repository import evidencias_count, listar_evidencias, registrar_artefacto
from src.repositories.parameter_repository import actualizar_parametro, obtener_parametro
from src.repositories.subject_repository import (
    get_asignatura_base,
    guardar_asignatura_base,
    limpiar_numero,
    listar_asignaturas_base,
)
from src.repositories.user_repository import (
    actualizar_usuario,
    autenticar_usuario,
    cambiar_password,
    crear_usuario,
    listar_usuarios,
)

# Servicios de negocio y exportadores
from src.services.academic_service import (
    analizar_coherencia_curso,
    expandir_plan_sesiones,
    generar_fechas_clase,
    generar_sugerencias_academicas,
    limpiar_df,
    matriz_riesgo_cursos,
    score_calidad_expediente,
    validar_plan,
)
from src.services.calendar_service import build_ics_calendar, parse_time_value
from src.services.docx_service import (
    add_paragraph_in_cell,
    agregar_parrafo_antes,
    aplicar_bordes_tabla,
    construir_analisis,
    crear_gc71_docx,
    crear_informe_ejecutivo_institucional_docx,
    crear_informe_gc72_docx,
    fijar_anchos_tabla,
    formato_entero,
    formato_porcentaje,
    normalizar_dataframe_gc72,
    porcentaje,
    remover_parrafo,
    set_cell_text,
    set_cell_width,
    set_label_value,
    set_section_text,
    shade_cell,
    texto_preformas,
)
from src.services.excel_service import (
    crear_plantilla_evaluacion_xlsx,
    crear_reporte_ejecutivo_xlsx,
    dataframe_to_xlsx_bytes,
    diagnostico_tabla_estudiantes,
    excel_bytes_sheets,
    leer_calificaciones,
    leer_excel_inteligente,
    leer_listado_estudiantes,
    normalizar_estudiantes_inteligente,
)
from src.services.export_service import (
    _hash_expediente,
    crear_paquete_curso_zip,
    nombre_archivo_seguro,
    safe_filename,
)
from src.services.health_service import health_status

# Interfaz UI y Router principal
from src.ui.router import main

if __name__ == "__main__":
    main()
