from __future__ import annotations

import os
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

APP_VERSION = "7.1.0-solid-enterprise"
DEFAULT_MAX_EVIDENCE_MB = 15

# Rutas del sistema
APP_DIR = Path(__file__).resolve().parents[1]
TEMPLATE_DIR = APP_DIR / "plantilla"
TEMPLATE_GC72 = TEMPLATE_DIR / "FD-GC72-Informe_Academico.docx"
TEMPLATE_GC71 = TEMPLATE_DIR / "FD-GC71.docx"
LOGO_POLI = TEMPLATE_DIR / "logo_poli.png"
LOGO_ICONTEC = TEMPLATE_DIR / "logo_icontec.png"
DATA_DIR = APP_DIR / "app_data"
DB_PATH = DATA_DIR / "fdgc_app.sqlite3"
EVIDENCE_DIR = DATA_DIR / "evidencias"
EXPORT_DIR = DATA_DIR / "exportaciones"
BACKUP_DIR = DATA_DIR / "backups"

# Módulos del sistema
MODULO_CENTRO = "Centro de control"
MODULO_EXPEDIENTE = "Expediente académico"
MODULO_PLANEADOR = "Planeador superior"
MODULO_EVIDENCIAS = "Evidencias y soportes"
MODULO_VALIDADOR = "Validador institucional"
MODULO_BACKUP = "Copias y restauración"
MODULO_DIAGNOSTICO = "Diagnóstico productivo"
MODULO_FLUJO = "Aprobaciones y versionamiento"
MODULO_MOTOR = "Motor académico"
MODULO_REPORTES = "Reportes ejecutivos"
MODULO_PARAMETROS = "Parámetros institucionales"
MODULO_BANCO = "Banco de asignaturas base"
MODULO_COHERENCIA = "Coherencia curricular"
MODULO_APROBACION_BLOQUEANTE = "Aprobación bloqueante"
MODULO_AUDITORIA_EXPEDIENTE = "Auditoría de expedientes"
MODULO_INFORME_INSTITUCIONAL = "Informe institucional"
MODULO_EXPORTACION_INSTITUCIONAL = "Paquetes masivos institucionales"
MODULO_CARGADOR = "Cargador inteligente"
MODULO_COMPARADOR = "Comparador de cortes"
MODULO_SEMAFORO = "Semáforo documental"
MODULO_EXPORTACION = "Exportación masiva"
MODULO_VERIFICACION = "Verificación de firmas"
MODULO_ASISTENTE = "Asistente académico IA"

# Mapeo de días de clase
DIAS_MAP: Dict[str, int] = {
    "Lunes": 0,
    "Martes": 1,
    "Miércoles": 2,
    "Miercoles": 2,
    "Jueves": 3,
    "Viernes": 4,
    "Sábado": 5,
    "Sabado": 5,
    "Domingo": 6,
}
DIAS_INV: Dict[int, str] = {v: k for k, v in DIAS_MAP.items() if "é" in k or k not in ("Miercoles", "Sabado")}

# Columnas y preformas FD-GC72
COLUMNAS_GC72 = [
    "Código",
    "Grupo",
    "Asignatura",
    "% Avance en contenido",
    "% Evaluado",
    "Estudiantes matriculados",
    "Desertaron N°",
    "Desertaron %",
    "Aprueban evaluación parcial N°",
    "Aprueban evaluación parcial %",
    "Reprueban evaluación parcial N°",
    "Reprueban evaluación parcial %",
    "Aprueban a la fecha N°",
    "Aprueban a la fecha %",
    "Reprueban a la fecha N°",
    "Reprueban a la fecha %",
]
NUMERICAS_ENTERAS_GC72 = [
    "Estudiantes matriculados",
    "Desertaron N°",
    "Aprueban evaluación parcial N°",
    "Reprueban evaluación parcial N°",
    "Aprueban a la fecha N°",
    "Reprueban a la fecha N°",
]
NUMERICAS_PORCENTAJE_GC72 = [
    "% Avance en contenido",
    "% Evaluado",
    "Desertaron %",
    "Aprueban evaluación parcial %",
    "Reprueban evaluación parcial %",
    "Aprueban a la fecha %",
    "Reprueban a la fecha %",
]
MAPA_TABLA_WORD_GC72 = COLUMNAS_GC72[:]

PREFORMAS_GC72: Dict[str, Dict[str, str]] = {
    "aspectos_positivos": {
        "Participación activa": "Se evidenció participación activa y disposición del grupo para el desarrollo de las actividades académicas propuestas.",
        "Avance adecuado": "El curso presentó un avance adecuado frente a los contenidos programados, manteniendo coherencia entre las actividades de clase y los resultados esperados.",
        "Aplicación práctica": "Los estudiantes lograron relacionar los conceptos abordados con situaciones prácticas, favoreciendo la comprensión aplicada de la asignatura.",
        "Mejora progresiva": "Se observó una mejora progresiva en la apropiación de los temas, especialmente en las actividades de seguimiento y retroalimentación.",
        "Buen cumplimiento": "La mayoría de los estudiantes cumplió con las actividades evaluativas y entregables definidos para el periodo reportado.",
        "Trabajo colaborativo": "El trabajo colaborativo fortaleció la discusión académica y permitió resolver dudas de manera más efectiva durante el curso.",
    },
    "inconvenientes": {
        "Inasistencia intermitente": "Se presentaron dificultades asociadas a inasistencia intermitente de algunos estudiantes, lo que afectó la continuidad del proceso formativo.",
        "Entregas tardías": "Algunos estudiantes realizaron entregas fuera de los tiempos establecidos, situación que limitó la retroalimentación oportuna.",
        "Brechas conceptuales": "Se identificaron brechas conceptuales en temas base, por lo cual fue necesario reforzar contenidos previos para avanzar con mayor solidez.",
        "Baja participación puntual": "En algunas sesiones se presentó baja participación de parte del grupo, especialmente en actividades que requerían preparación previa.",
        "Dificultades técnicas": "Se presentaron dificultades técnicas o de acceso a herramientas requeridas para el desarrollo de algunas actividades académicas.",
        "Carga académica acumulada": "La acumulación de actividades de otros espacios académicos incidió en el ritmo de trabajo y en la oportunidad de algunas entregas.",
    },
    "propuestas": {
        "Seguimiento formativo": "Fortalecer el seguimiento formativo mediante actividades cortas de verificación, retroalimentación temprana y acompañamiento focalizado.",
        "Talleres aplicados": "Incorporar talleres aplicados por unidades temáticas para consolidar la relación entre teoría, práctica y evaluación.",
        "Nivelación inicial": "Implementar una actividad diagnóstica y espacios de nivelación para atender brechas conceptuales antes de abordar contenidos de mayor complejidad.",
        "Rúbricas claras": "Socializar rúbricas y criterios de evaluación desde el inicio de cada actividad, con el fin de mejorar la calidad de las entregas.",
        "Aprendizaje basado en problemas": "Utilizar ejercicios basados en problemas reales del contexto profesional para incrementar la pertinencia y motivación del proceso formativo.",
        "Alertas tempranas": "Aplicar alertas tempranas frente a inasistencia, bajo desempeño o entregas pendientes, articulando acciones de mejora con los estudiantes.",
    },
}

# Columnas FD-GC71
COLUMNAS_MODULOS = [
    "Unidad",
    "Contenido / tema central",
    "Horas presenciales",
    "Sesiones",
    "Trabajo presencial",
    "Trabajo independiente",
]
COLUMNAS_HORARIOS = ["Día", "Hora inicio", "Hora fin", "Lugar / ambiente"]
COLUMNAS_EVALUACIONES = [
    "Tipo de evaluación",
    "Procedimiento de evaluación",
    "Valor (%)",
    "Fecha de realización",
    "Unidad relacionada",
    "Corte",
]
COLUMNAS_SESIONES = [
    "Unidad",
    "N° sesión",
    "Fecha",
    "Horario",
    "Contenido por desarrollar",
    "Descripción del trabajo presencial",
    "Descripción trabajo independiente",
    "Lugar / ambiente",
]

TEXTOS_PREDEFINIDOS_GC71 = {
    "justificacion": "La asignatura aporta a la formación académica y profesional mediante la articulación entre fundamentos conceptuales, aplicación práctica y análisis de situaciones propias del campo disciplinar. Su desarrollo favorece la comprensión de problemas del contexto, el uso de herramientas técnicas y la toma de decisiones sustentada en criterios académicos, éticos y profesionales.",
    "competencias": "La asignatura tributa al desarrollo de competencias asociadas con el análisis crítico, la resolución de problemas, la comunicación técnica, el trabajo colaborativo y la aplicación de conocimientos disciplinares en escenarios reales o simulados.",
    "resultados": "Al finalizar la asignatura, el estudiante estará en capacidad de reconocer, interpretar y aplicar los conceptos y procedimientos centrales del curso, integrando evidencias, criterios técnicos y estrategias de solución acordes con los resultados de aprendizaje del programa.",
    "objetivo_general": "Desarrollar en el estudiante capacidades conceptuales, metodológicas y prácticas para comprender y aplicar los contenidos de la asignatura en situaciones propias de su formación académica y profesional.",
    "objetivos_especificos": "1. Reconocer los fundamentos conceptuales de la asignatura.\n2. Aplicar procedimientos y herramientas propias del área de formación.\n3. Analizar casos, ejercicios o problemas relacionados con el contexto profesional.\n4. Comunicar resultados de manera clara, ordenada y técnicamente sustentada.",
    "metodologias": "La asignatura se desarrollará mediante clases orientadoras, talleres aplicados, análisis de casos, ejercicios prácticos, aprendizaje basado en problemas, socialización de avances y retroalimentación permanente. Se promoverá la participación activa del estudiante y la integración entre trabajo presencial e independiente.",
    "ambientes": "Aula de clase, plataforma institucional, recursos digitales de apoyo y, cuando aplique, laboratorios, salidas pedagógicas o ambientes especializados requeridos para el logro de los resultados de aprendizaje.",
    "medios": "Presentaciones, guías de clase, material bibliográfico, bases de datos académicas, plataforma virtual, software especializado cuando aplique, tablero, equipos audiovisuales y recursos institucionales necesarios para el desarrollo de las actividades.",
    "referencias": "Bibliografía básica y complementaria definida por el docente, documentos institucionales del programa, artículos académicos recientes, recursos digitales especializados y fuentes en segunda lengua cuando sean pertinentes para la asignatura.",
}


def _get_secret_value(name: str, default: Optional[str] = None) -> Optional[str]:
    """Lee configuración desde Streamlit secrets o variables de entorno sin romper ejecución local."""
    try:
        import streamlit as _st
        if hasattr(_st, "secrets") and name in _st.secrets:
            return _st.secrets.get(name)
    except Exception:
        pass
    return os.getenv(name, default)


def get_app_env() -> str:
    return str(_get_secret_value("APP_ENV", os.getenv("APP_ENV", "local")) or "local").lower().strip()


def get_database_url() -> str:
    return str(_get_secret_value("DATABASE_URL", os.getenv("DATABASE_URL", "")) or "").strip()


def _database_url_es_placeholder(url: str) -> bool:
    if not url:
        return False
    u = url.upper()
    marcadores = ["USUARIO", "CLAVE", "HOST", "BASE", "PASSWORD", "XXXXX"]
    return any(m in u for m in marcadores)


def usar_postgres() -> bool:
    url = get_database_url().strip()
    low = url.lower()
    if _database_url_es_placeholder(url):
        return False
    return low.startswith("postgres://") or low.startswith("postgresql://")


def postgres_url_normalizada() -> str:
    url = get_database_url()
    if url.startswith("postgres://"):
        return "postgresql://" + url[len("postgres://"):]
    return url


def _safe_int_secret(name: str, default: int) -> int:
    raw = _get_secret_value(name, str(default))
    try:
        return int(str(raw).strip())
    except Exception:
        return default


def max_evidence_bytes() -> int:
    mb = _safe_int_secret("MAX_EVIDENCE_MB", DEFAULT_MAX_EVIDENCE_MB)
    return mb * 1024 * 1024


def initial_admin_config() -> Tuple[str, str, str, str]:
    """Credenciales de arranque parametrizables por secretos."""
    return (
        str(_get_secret_value("INITIAL_ADMIN_USER", "admin") or "admin"),
        str(_get_secret_value("INITIAL_ADMIN_PASSWORD", "Admin123*") or "Admin123*"),
        str(_get_secret_value("INITIAL_ADMIN_NAME", "Administrador del sistema") or "Administrador del sistema"),
        str(_get_secret_value("INITIAL_ADMIN_EMAIL", "") or ""),
    )
