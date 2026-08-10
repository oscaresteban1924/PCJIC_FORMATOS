from __future__ import annotations

from typing import Any, Dict, List, Optional

from src.config import (
    MODULO_APROBACION_BLOQUEANTE,
    MODULO_ASISTENTE,
    MODULO_AUDITORIA_EXPEDIENTE,
    MODULO_BACKUP,
    MODULO_BANCO,
    MODULO_CARGADOR,
    MODULO_CENTRO,
    MODULO_COHERENCIA,
    MODULO_COMPARADOR,
    MODULO_DIAGNOSTICO,
    MODULO_EVIDENCIAS,
    MODULO_EXPEDIENTE,
    MODULO_EXPORTACION,
    MODULO_EXPORTACION_INSTITUCIONAL,
    MODULO_FLUJO,
    MODULO_INFORME_INSTITUCIONAL,
    MODULO_MOTOR,
    MODULO_PARAMETROS,
    MODULO_PLANEADOR,
    MODULO_REPORTES,
    MODULO_SEMAFORO,
    MODULO_VALIDADOR,
    MODULO_VERIFICACION,
)
from src.security import ROLES_PERMISOS

UX_MODULE_META: Dict[str, Dict[str, Any]] = {
    "Inicio": {"icon": "🏠", "group": "Inicio", "desc": "Resumen institucional y acceso rápido.", "step": 1},
    MODULO_BANCO: {"icon": "🏛️", "group": "Dirección", "desc": "Banco institucional de microcurrículos y asignaturas.", "step": 2},
    MODULO_COHERENCIA: {"icon": "⚖️", "group": "Dirección", "desc": "Auditoría de coherencia curricular y créditos.", "step": 2},
    MODULO_APROBACION_BLOQUEANTE: {"icon": "🔒", "group": "Gobierno", "desc": "Cierre institucional con bloqueo contra ediciones.", "step": 5},
    MODULO_INFORME_INSTITUCIONAL: {"icon": "📜", "group": "Gobierno", "desc": "Informe consolidado institucional en Word/PDF.", "step": 6},
    MODULO_EXPORTACION_INSTITUCIONAL: {"icon": "📦", "group": "Gobierno", "desc": "Paquete ZIP masivo de programas y facultades.", "step": 7},
    MODULO_AUDITORIA_EXPEDIENTE: {"icon": "🔎", "group": "Gobierno", "desc": "Trazabilidad completa por versión y eventos.", "step": 6},
    MODULO_CENTRO: {"icon": "📊", "group": "Dirección", "desc": "Indicadores, riesgo y salud académica.", "step": 2},
    MODULO_EXPEDIENTE: {"icon": "📚", "group": "Operación", "desc": "Carpeta viva del curso y trazabilidad.", "step": 3},
    MODULO_PLANEADOR: {"icon": "🧭", "group": "Operación", "desc": "Planeación automática, calendario y FD-GC71.", "step": 4},
    MODULO_FLUJO: {"icon": "✅", "group": "Gobierno", "desc": "Revisión, observaciones, aprobación y cierre.", "step": 5},
    MODULO_MOTOR: {"icon": "⚙️", "group": "Operación", "desc": "Motor de cálculo académico e indicadores.", "step": 4},
    MODULO_REPORTES: {"icon": "📈", "group": "Dirección", "desc": "Reportes ejecutivos descargables.", "step": 6},
    MODULO_PARAMETROS: {"icon": "🛠️", "group": "Administración", "desc": "Parámetros institucionales.", "step": 8},
    "FD-GC71 - Planeación": {"icon": "📝", "group": "Formatos", "desc": "Guía didáctica y concertación evaluativa.", "step": 4},
    "FD-GC72 - Informe académico": {"icon": "📋", "group": "Formatos", "desc": "Informe académico y análisis cualitativo.", "step": 6},
    MODULO_EVIDENCIAS: {"icon": "📎", "group": "Operación", "desc": "Gestión de evidencias y soportes firmados.", "step": 5},
    MODULO_VALIDADOR: {"icon": "🛡️", "group": "Gobierno", "desc": "Validador documental institucional.", "step": 6},
    MODULO_CARGADOR: {"icon": "📥", "group": "Operación", "desc": "Normalizador de listados y notas Excel.", "step": 3},
    MODULO_COMPARADOR: {"icon": "🔀", "group": "Operación", "desc": "Comparador de cortes e indicadores.", "step": 6},
    MODULO_SEMAFORO: {"icon": "🚥", "group": "Gobierno", "desc": "Semáforo documental de riesgo.", "step": 5},
    MODULO_EXPORTACION: {"icon": "🗂️", "group": "Gobierno", "desc": "Exportación masiva de paquetes.", "step": 7},
    MODULO_VERIFICACION: {"icon": "🔑", "group": "Gobierno", "desc": "Verificación criptográfica de expedientes.", "step": 8},
    MODULO_ASISTENTE: {"icon": "🤖", "group": "Operación", "desc": "Asistente académico IA.", "step": 4},
    MODULO_BACKUP: {"icon": "💾", "group": "Sistema", "desc": "Copias de seguridad y restauración.", "step": 8},
    MODULO_DIAGNOSTICO: {"icon": "🩺", "group": "Sistema", "desc": "Diagnóstico de infraestructura.", "step": 8},
    "Usuarios y perfiles": {"icon": "👥", "group": "Administración", "desc": "Control de usuarios, roles y accesos.", "step": 8},
    "Auditoría": {"icon": "📜", "group": "Sistema", "desc": "Bitácora de seguridad y auditoría.", "step": 8},
    "Mi cuenta": {"icon": "👤", "group": "Cuenta", "desc": "Configuración personal y clave.", "step": 8},
    "Ayuda / flujo recomendado": {"icon": "❓", "group": "Ayuda", "desc": "Flujo de trabajo y preguntas frecuentes.", "step": 1},
}


def ux_meta(modulo: str) -> Dict[str, Any]:
    return UX_MODULE_META.get(modulo, {"icon": "▫️", "group": "Otros", "desc": "Módulo del sistema.", "step": 1})


def ux_label(modulo: str) -> str:
    meta = ux_meta(modulo)
    return f"{meta.get('icon', '▫️')} {modulo}"


def ux_apply_theme(st: Any):
    """Inyecta el sistema visual premium con legibilidad perfecta."""
    st.markdown(
        """
<style>
:root {
    --ux-bg: #f6f8fc;
    --ux-card: #ffffff;
    --ux-ink: #152238;
    --ux-muted: #607089;
    --ux-line: rgba(21,34,56,.11);
    --ux-primary: #1f4fd8;
    --ux-primary-2: #0f2f86;
    --ux-accent: #14b8a6;
    --ux-shadow: 0 12px 32px rgba(15, 31, 72, .08);
    --ux-radius: 22px;
}
html, body, [data-testid="stAppViewContainer"] {
    background:
        radial-gradient(circle at top left, rgba(31,79,216,.10), transparent 35%),
        radial-gradient(circle at top right, rgba(20,184,166,.10), transparent 30%),
        var(--ux-bg) !important;
}
.block-container {
    padding-top: 1.15rem !important;
    padding-bottom: 4rem !important;
    max-width: 1480px !important;
}
[data-testid="stHeader"] { background: rgba(246,248,252,.72) !important; backdrop-filter: blur(10px); }

/* Sidebar styling */
[data-testid="stSidebar"] > div:first-child {
    background: linear-gradient(180deg, #0f172a 0%, #15254c 52%, #0f172a 100%) !important;
    color: #ffffff !important;
    border-right: 1px solid rgba(255,255,255,.08);
}
[data-testid="stSidebar"] label, 
[data-testid="stSidebar"] .stMarkdown,
[data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {
    color: #ffffff !important;
}
[data-testid="stSidebar"] [data-testid="stCaptionContainer"] {
    color: rgba(255,255,255,0.75) !important;
}

/* BaseWeb Selectbox styling fix - High Contrast & Crisp Visibility */
[data-testid="stSidebar"] [data-baseweb="select"] {
    background-color: #ffffff !important;
    border-radius: 12px !important;
    border: 1px solid rgba(255,255,255,0.3) !important;
}
[data-testid="stSidebar"] [data-baseweb="select"] * {
    color: #0f172a !important;
    font-weight: 600 !important;
}
div[data-baseweb="popover"] * {
    color: #0f172a !important;
    background-color: #ffffff !important;
}

/* Sidebar Button styling fix - Ensures Cerrar sesión button text is 100% visible */
[data-testid="stSidebar"] .stButton > button,
[data-testid="stSidebar"] .stDownloadButton > button {
    background-color: #ffffff !important;
    color: #0f172a !important;
    border: 1px solid rgba(15, 23, 42, 0.2) !important;
    border-radius: 12px !important;
    font-weight: 700 !important;
}
[data-testid="stSidebar"] .stButton > button *,
[data-testid="stSidebar"] .stDownloadButton > button * {
    color: #0f172a !important;
}
[data-testid="stSidebar"] .stButton > button:hover,
[data-testid="stSidebar"] .stDownloadButton > button:hover {
    background-color: #dc2626 !important;
    color: #ffffff !important;
    border-color: #dc2626 !important;
}
[data-testid="stSidebar"] .stButton > button:hover *,
[data-testid="stSidebar"] .stDownloadButton > button:hover * {
    color: #ffffff !important;
}

[data-testid="stMetric"] {
    background: rgba(255,255,255,.88);
    border: 1px solid var(--ux-line);
    padding: 1rem 1.05rem;
    border-radius: 18px;
    box-shadow: 0 10px 28px rgba(15,31,72,.06);
}
.stButton > button, .stDownloadButton > button {
    border-radius: 14px !important;
    border: 1px solid rgba(31,79,216,.18) !important;
    font-weight: 700 !important;
    min-height: 2.65rem;
}
.ux-hero {
    border-radius: 28px;
    padding: 1.35rem 1.45rem;
    background: linear-gradient(135deg, rgba(31,79,216,.96), rgba(15,47,134,.96) 58%, rgba(20,184,166,.88));
    color: #fff;
    box-shadow: 0 20px 48px rgba(15,47,134,.22);
    margin-bottom: 1.05rem;
    border: 1px solid rgba(255,255,255,.18);
}
.ux-hero h1 { margin: 0 0 .35rem 0; font-size: 1.75rem; line-height: 1.15; color: #ffffff !important; }
.ux-hero p { margin: 0; color: rgba(255,255,255,.86) !important; font-size: .98rem; }
.ux-card {
    background: rgba(255,255,255,.92);
    border: 1px solid var(--ux-line);
    border-radius: var(--ux-radius);
    padding: 1rem;
    box-shadow: var(--ux-shadow);
}
.ux-footer-note { text-align:center; color: var(--ux-muted); font-size:.82rem; margin-top:2.5rem; }
</style>
        """,
        unsafe_allow_html=True,
    )


def ux_badge(label: str, kind: str = "") -> str:
    bg = "rgba(31,79,216,.10)"
    color = "#1f4fd8"
    if kind == "success":
        bg = "rgba(22,163,74,.12)"
        color = "#15803d"
    elif kind == "warning":
        bg = "rgba(245,158,11,.14)"
        color = "#b45309"
    elif kind == "danger":
        bg = "rgba(239,68,68,.14)"
        color = "#b91c1c"
    return f"<span style='display:inline-block; padding:.18rem .55rem; border-radius:999px; background:{bg}; color:{color}; font-weight:800; font-size:.78rem;'>{label}</span>"


def ux_card(icon: str, title: str, text: str) -> str:
    return f"""
    <div class="ux-card">
        <div class="ux-card-icon">{icon}</div>
        <h3>{title}</h3>
        <p>{text}</p>
    </div>
    """


def ux_render_hero(st: Any, modulo: str, user: Dict[str, Any]):
    meta = ux_meta(modulo)
    st.markdown(
        f"""
    <div class="ux-hero">
        <div style="display:flex; justify-content:space-between; align-items:flex-start;">
            <div>
                <h1>{meta.get('icon','🎓')} {modulo}</h1>
                <p>{meta.get('desc','Gestión académica institucional.')}</p>
            </div>
            <div style="text-align:right;">
                <div style="background:rgba(255,255,255,0.2); padding:0.25rem 0.75rem; border-radius:12px; font-weight:600;">👤 {user.get('nombre_completo', user.get('usuario',''))}</div>
                <div style="background:rgba(255,255,255,0.2); padding:0.25rem 0.75rem; border-radius:12px; font-weight:600; margin-top:.35rem;">🛡️ Perfil: {user.get('rol','Docente')}</div>
            </div>
        </div>
    </div>
    """,
        unsafe_allow_html=True,
    )


def ux_render_path(st: Any, active_module: str):
    pass


def ux_sidebar(st: Any, user: Dict[str, Any]) -> str:
    with st.sidebar:
        st.markdown("### 🎓 Gestor Académico")
        st.caption(f"Usuario: {user.get('nombre_completo', user.get('usuario',''))}")
        modulos = ROLES_PERMISOS.get(user.get("rol"), {}).get("modulos", ["Inicio", "Mi cuenta", "Ayuda / flujo recomendado"])

        seleccion = st.selectbox(
            "Seleccione Módulo",
            options=modulos,
            format_func=ux_label,
            key="sidebar_module_select",
        )

        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🚪 Cerrar sesión", use_container_width=True):
            st.session_state.clear()
            st.rerun()

        return seleccion
