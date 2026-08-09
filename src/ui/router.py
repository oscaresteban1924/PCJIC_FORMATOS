from __future__ import annotations

import streamlit as st

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
from src.database.schema import init_db
from src.security import tiene_permiso
from src.ui.auth_views import pantalla_login, ui_admin_usuarios, ui_mi_cuenta
from src.ui.course_views import (
    ui_cargador_inteligente,
    ui_comparador_cortes,
    ui_evidencias,
    ui_expediente_academico,
    ui_gc71,
    ui_gc72,
    ui_planeador_superior,
    ui_semaforo_expediente,
)
from src.ui.institutional_views import (
    ui_aprobacion_bloqueante,
    ui_asistente_academico,
    ui_auditoria,
    ui_auditoria_expediente,
    ui_ayuda,
    ui_backup,
    ui_banco_asignaturas,
    ui_centro_control,
    ui_coherencia_academica,
    ui_diagnostico_productivo,
    ui_exportacion_institucional,
    ui_exportacion_masiva,
    ui_flujo_aprobaciones,
    ui_informe_institucional,
    ui_inicio,
    ui_motor_academico,
    ui_parametros,
    ui_reportes_ejecutivos,
    ui_verificacion_documental,
)
from src.ui.theme import ux_apply_theme, ux_render_hero, ux_render_path, ux_sidebar


def main():
    st.set_page_config(
        page_title="Gestor Académico Institucional FD-GC71 / FD-GC72",
        page_icon="🎓",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    ux_apply_theme(st)
    try:
        init_db()
    except Exception as exc:
        st.error("La aplicación inició, pero no pudo preparar la base de datos.")
        st.markdown(
            """
        Esto suele pasar cuando `DATABASE_URL` está mal escrita,
        la base de datos externa no acepta conexiones o se especificó un puerto incorrecto.
        """
        )
        st.code(str(exc), language="text")
        st.stop()

    if "auth_user" not in st.session_state:
        pantalla_login(st)
        return

    user = st.session_state.get("auth_user", {})
    pagina = ux_sidebar(st, user)

    if not tiene_permiso(pagina, user):
        st.error("Este perfil no tiene permisos para abrir este módulo.")
        return

    ux_render_hero(st, pagina, user)
    if pagina not in ["Inicio", "Mi cuenta", "Ayuda / flujo recomendado", MODULO_DIAGNOSTICO]:
        ux_render_path(st, pagina)

    if pagina == "Inicio":
        ui_inicio(st)
    elif pagina == MODULO_BANCO:
        ui_banco_asignaturas(st)
    elif pagina == MODULO_COHERENCIA:
        ui_coherencia_academica(st)
    elif pagina == MODULO_APROBACION_BLOQUEANTE:
        ui_aprobacion_bloqueante(st)
    elif pagina == MODULO_INFORME_INSTITUCIONAL:
        ui_informe_institucional(st)
    elif pagina == MODULO_EXPORTACION_INSTITUCIONAL:
        ui_exportacion_institucional(st)
    elif pagina == MODULO_AUDITORIA_EXPEDIENTE:
        ui_auditoria_expediente(st)
    elif pagina == MODULO_CENTRO:
        ui_centro_control(st)
    elif pagina == MODULO_SEMAFORO:
        ui_semaforo_expediente(st)
    elif pagina == MODULO_EXPEDIENTE:
        ui_expediente_academico(st)
    elif pagina == MODULO_PLANEADOR:
        ui_planeador_superior(st)
    elif pagina == MODULO_ASISTENTE:
        ui_asistente_academico(st)
    elif pagina == MODULO_CARGADOR:
        ui_cargador_inteligente(st)
    elif pagina == MODULO_COMPARADOR:
        ui_comparador_cortes(st)
    elif pagina == MODULO_FLUJO:
        ui_flujo_aprobaciones(st)
    elif pagina == MODULO_MOTOR:
        ui_motor_academico(st)
    elif pagina == MODULO_REPORTES:
        ui_reportes_ejecutivos(st)
    elif pagina == MODULO_EXPORTACION:
        ui_exportacion_masiva(st)
    elif pagina == MODULO_VERIFICACION:
        ui_verificacion_documental(st)
    elif pagina == MODULO_PARAMETROS:
        ui_parametros(st)
    elif pagina.startswith("FD-GC71"):
        ui_gc71(st)
    elif pagina.startswith("FD-GC72"):
        ui_gc72(st)
    elif pagina == MODULO_EVIDENCIAS:
        ui_evidencias(st)
    elif pagina == MODULO_VALIDADOR:
        ui_inicio(st)
    elif pagina == MODULO_BACKUP:
        ui_backup(st)
    elif pagina == MODULO_DIAGNOSTICO:
        ui_diagnostico_productivo(st)
    elif pagina == "Usuarios y perfiles":
        ui_admin_usuarios(st)
    elif pagina == "Auditoría":
        ui_auditoria(st)
    elif pagina == "Mi cuenta":
        ui_mi_cuenta(st)
    else:
        ui_ayuda(st)

    st.markdown("<div class='ux-footer-note'>Gestor Académico Institucional · banco → planeación → coherencia → aprobación → cortes → informe → exportación.</div>", unsafe_allow_html=True)


if __name__ == "__main__":
    main()
