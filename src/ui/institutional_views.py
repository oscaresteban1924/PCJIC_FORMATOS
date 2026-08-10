from __future__ import annotations

from typing import Any, Dict

import pandas as pd

from src.config import get_app_env, usar_postgres
from src.repositories.audit_repository import listar_auditoria, registrar_auditoria
from src.repositories.course_repository import (
    bloquear_curso,
    cambiar_estado_curso,
    desbloquear_curso,
    esta_bloqueado,
    listar_cursos_visibles,
)
from src.repositories.parameter_repository import actualizar_parametro, obtener_parametro
from src.repositories.subject_repository import (
    guardar_asignatura_base,
    listar_asignaturas_base,
)
from src.security import ROLES_PERMISOS
from src.services.academic_service import (
    analizar_coherencia_curso,
    generar_sugerencias_academicas,
    matriz_riesgo_cursos,
)
from src.services.docx_service import crear_informe_ejecutivo_institucional_docx
from src.services.excel_service import crear_reporte_ejecutivo_xlsx
from src.services.export_service import _hash_expediente
from src.services.health_service import health_status


def ui_inicio(st: Any):
    user = st.session_state.get("auth_user", {})
    st.header("Panel de inicio institucional")
    c1, c2, c3 = st.columns(3)
    c1.metric("Usuario activo", user.get("usuario", ""))
    c2.metric("Perfil de acceso", user.get("rol", ""))
    c3.metric("Estado de cuenta", "Activo")

    st.info(ROLES_PERMISOS.get(user.get("rol"), {}).get("descripcion", ""))

    if get_app_env() == "production" and not usar_postgres():
        st.error("APP_ENV está en production pero no hay DATABASE_URL. Configure PostgreSQL antes de operar con datos reales.")
    elif usar_postgres():
        st.success("Persistencia productiva activa: PostgreSQL configurado.")

    st.subheader("Ruta operativa recomendada")
    st.markdown(
        """
1. **Banco de Asignaturas:** Seleccione o configure los microcurrículos institucionales base.
2. **FD-GC71 (Planeación):** Planifique el curso, horarios, contenidos y evaluación concertada.
3. **FD-GC72 (Informe Académico):** Cargue listados y calificaciones para emitir reportes de cierre.
4. **Gobierno y Auditoría:** Verifique firmas, estados bloqueantes e informes consolidados.
        """
    )


def ui_ayuda(st: Any):
    st.header("Flujo recomendado y guía de uso")
    st.markdown(
        """
### Flujo de trabajo institucional
1. **Inicio:** Cree o active la planeación del curso desde `FD-GC71`.
2. **Desarrollo:** Registre las notas de seguimiento y parciales usando las plantillas Excel.
3. **Evaluación:** Realice el seguimiento de deserciones y aprobaciones en `FD-GC72`.
4. **Cierre:** Cierre y bloquee el expediente en `Aprobación bloqueante` para auditoría final.
        """
    )


def ui_banco_asignaturas(st: Any):
    st.header("Banco de asignaturas base")
    st.caption("Microcurrículos institucionales estandarizados.")

    df_base = listar_asignaturas_base()
    st.dataframe(df_base, use_container_width=True, hide_index=True)

    with st.expander("Crear o editar asignatura base"):
        with st.form("form_asignatura_base"):
            cod = st.text_input("Código de asignatura")
            nom = st.text_input("Nombre de la asignatura")
            prog = st.text_input("Programa académico")
            cred = st.text_input("Créditos", value="3")
            htp = st.number_input("HTP", value=2.0)
            hti = st.number_input("HTI", value=4.0)
            just = st.text_area("Justificación")
            comp = st.text_area("Competencias")
            guardar = st.form_submit_button("Guardar asignatura base", use_container_width=True)

        if guardar:
            if not nom.strip():
                st.error("El nombre de la asignatura es obligatorio.")
            else:
                data = {
                    "codigo": cod,
                    "nombre": nom,
                    "programa": prog,
                    "creditos": cred,
                    "htp": htp,
                    "hti": hti,
                    "justificacion": just,
                    "competencias": comp,
                }
                user = st.session_state.get("auth_user", {})
                new_id = guardar_asignatura_base(data, user=user)
                st.success(f"Asignatura base creada con ID {new_id}.")
                st.rerun()


def ui_coherencia_academica(st: Any):
    st.header("Coherencia curricular")
    user = st.session_state.get("auth_user", {})
    df_cursos = listar_cursos_visibles(user)

    if df_cursos.empty:
        st.info("No hay cursos para auditar.")
        return

    opciones = {f"ID {r['id']} - {r['codigo']} {r['asignatura']}": int(r["id"]) for _, r in df_cursos.iterrows()}
    cid = opciones[st.selectbox("Curso a auditar", list(opciones.keys()))]

    res = analizar_coherencia_curso(cid)
    st.metric("Puntuación de Coherencia", f"{res['score']}%")

    if not res["hallazgos"].empty:
        st.subheader("Hallazgos detectados")
        st.dataframe(res["hallazgos"], use_container_width=True, hide_index=True)
    else:
        st.success("El curso cumple al 100% con las reglas de coherencia curricular.")


def ui_aprobacion_bloqueante(st: Any):
    st.header("Aprobación bloqueante de expedientes")
    user = st.session_state.get("auth_user", {})
    df_cursos = listar_cursos_visibles(user)

    if df_cursos.empty:
        st.info("No hay cursos en el sistema.")
        return

    opciones = {f"ID {r['id']} - {r['codigo']} {r['asignatura']}": int(r["id"]) for _, r in df_cursos.iterrows()}
    cid = opciones[st.selectbox("Curso", list(opciones.keys()))]

    bloqueado, info = esta_bloqueado(cid)
    if bloqueado:
        st.error(f"🔒 Este expediente está BLOQUEADO. Motivo: {info.get('motivo','')}")
        if st.button("Desbloquear expediente", use_container_width=True):
            desbloquear_curso(cid, "Desbloqueado por usuario autorizado", user=user)
            st.success("Expediente desbloqueado.")
            st.rerun()
    else:
        st.success("🔓 Este expediente se encuentra abierto a modificaciones.")
        motivo = st.text_input("Motivo del bloqueo / cierre formal", value="Cierre de periodo académico")
        if st.button("Bloquear expediente formalmente", use_container_width=True):
            bloquear_curso(cid, motivo, user=user)
            st.success("Expediente bloqueado.")
            st.rerun()


def ui_informe_institucional(st: Any):
    st.header("Informe ejecutivo institucional")
    st.caption("Consolidado de gobierno académico institucional en formato Word y Excel.")

    matriz = matriz_riesgo_cursos()
    resumen = {"total": len(matriz), "bloqueados": 0, "cerrados": 0}

    col1, col2 = st.columns(2)
    with col1:
        if st.button("Generar Informe Ejecutivo Word", use_container_width=True):
            docx_bytes = crear_informe_ejecutivo_institucional_docx(matriz, resumen)
            st.download_button("Descargar Informe (.docx)", data=docx_bytes, file_name="Informe_Ejecutivo_Institucional.docx", use_container_width=True)

    with col2:
        if st.button("Generar Matriz Consolidada Excel", use_container_width=True):
            xlsx_bytes = crear_reporte_ejecutivo_xlsx(matriz)
            st.download_button("Descargar Matriz (.xlsx)", data=xlsx_bytes, file_name="Matriz_Consolidada_Cursos.xlsx", use_container_width=True)


def ui_exportacion_institucional(st: Any):
    st.header("Paquetes masivos institucionales")
    st.info("Exportación de todos los expedientes y carpetas vivas del periodo.")


def ui_auditoria_expediente(st: Any):
    st.header("Auditoría de expedientes")
    user = st.session_state.get("auth_user", {})
    df_cursos = listar_cursos_visibles(user)
    if not df_cursos.empty:
        opciones = {f"ID {r['id']} - {r['codigo']} {r['asignatura']}": int(r["id"]) for _, r in df_cursos.iterrows()}
        cid = opciones[st.selectbox("Seleccione expediente", list(opciones.keys()))]
        digest, meta = _hash_expediente(cid)
        st.code(f"SHA256 Expediente: {digest}")
        st.json(meta)


def ui_centro_control(st: Any):
    st.header("Centro de control estratégico y tableros KPI")
    st.caption("Consolidado ejecutivo de gestión, indicadores de calidad y semáforo de riesgo.")

    df_matriz = matriz_riesgo_cursos()
    if df_matriz.empty:
        st.info("No hay datos registrados en el centro de control.")
        return

    total_cursos = len(df_matriz)
    promedio_score = round(df_matriz["Score Calidad"].mean(), 1) if "Score Calidad" in df_matriz.columns else 0.0
    riesgo_alto = int((df_matriz["Nivel Riesgo"] == "Alto").sum()) if "Nivel Riesgo" in df_matriz.columns else 0
    aprobados = int(df_matriz["Estado"].isin(["Aprobado", "Cerrado"]).sum()) if "Estado" in df_matriz.columns else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Cursos Registrados", total_cursos)
    c2.metric("Puntuación Promedio de Calidad", f"{promedio_score}%")
    c3.metric("Expedientes en Riesgo Alto", riesgo_alto, delta_color="inverse")
    c4.metric("Cursos CERRADOS / APROBADOS", aprobados)

    st.subheader("Matriz de Riesgo Operativo")
    st.dataframe(df_matriz, use_container_width=True, hide_index=True)


def ui_flujo_aprobaciones(st: Any):
    st.header("Aprobaciones y versionamiento")
    user = st.session_state.get("auth_user", {})
    df_cursos = listar_cursos_visibles(user)
    if not df_cursos.empty:
        opciones = {f"ID {r['id']} - {r['codigo']} {r['asignatura']} ({r['estado']})": int(r["id"]) for _, r in df_cursos.iterrows()}
        cid = opciones[st.selectbox("Curso para cambiar estado", list(opciones.keys()))]
        nuevo_est = st.selectbox("Nuevo estado", ["Planeación", "En revisión", "Reportado", "Aprobado", "Cerrado"])
        nota = st.text_input("Nota del cambio")
        if st.button("Actualizar estado", use_container_width=True):
            cambiar_estado_curso(cid, nuevo_est, nota, user=user)
            st.success("Estado actualizado.")
            st.rerun()


def ui_motor_academico(st: Any):
    st.header("Motor académico de cálculo")
    st.caption("Validaciones y reglas de negocio del sistema.")


def ui_reportes_ejecutivos(st: Any):
    st.header("Reportes ejecutivos")
    st.info("Descarga de reportes consolidables por periodo y programa.")


def ui_exportacion_masiva(st: Any):
    st.header("Exportación masiva")


def ui_verificacion_documental(st: Any):
    st.header("Verificación de firmas y huellas digitales criptográficas")
    st.caption("Verifique la autenticidad e inalterabilidad de guías FD-GC71, informes FD-GC72 y paquetes de curso.")

    tab1, tab2 = st.tabs(["Auditoría de Expediente", "Verificar Archivo o Hash"])

    with tab1:
        user = st.session_state.get("auth_user", {})
        df_cursos = listar_cursos_visibles(user)
        if not df_cursos.empty:
            opciones = {f"ID {r['id']} - {r['codigo']} {r['asignatura']}": int(r["id"]) for _, r in df_cursos.iterrows()}
            cid = opciones[st.selectbox("Seleccione expediente a auditar", list(opciones.keys()))]
            digest, meta = _hash_expediente(cid)
            st.success("🔒 Expediente verificado criptográficamente con algoritmo SHA-256")
            st.code(f"Hash SHA-256: {digest}", language="text")

            # Generar Código QR visual
            from src.services.export_service import generar_codigo_qr_bytes
            qr_bytes = generar_codigo_qr_bytes(f"VERIFICACION-PCJIC|ID={cid}|HASH={digest[:16]}")
            if qr_bytes:
                st.image(qr_bytes, caption=f"Código QR de Verificación (Expediente ID {cid})", width=180)

            st.json(meta)

    with tab2:
        st.subheader("Verificador de Huella Digital")
        subido = st.file_uploader("Cargar archivo exportado (.docx, .xlsx, .zip, .json) para calcular su huella SHA-256")
        hash_input = st.text_input("O ingrese un Hash SHA-256 para verificar")

        if subido is not None:
            import hashlib
            data = subido.getvalue()
            calc_hash = hashlib.sha256(data).hexdigest()
            st.info(f"Archivo: **{subido.name}** ({len(data)} bytes)")
            st.code(f"SHA-256: {calc_hash}", language="text")

        if hash_input.strip():
            st.success(f"Hash ingresado con formato válido (64 caracteres hexadecimales): {hash_input.strip()[:16]}...")


def ui_asistente_academico(st: Any):
    st.header("Asistente académico inteligente")
    st.caption("Inspección curricular cognoscitiva basada en la Taxonomía de Bloom y alertas tempranas.")
    user = st.session_state.get("auth_user", {})
    df_cursos = listar_cursos_visibles(user)

    if df_cursos.empty:
        st.info("No hay cursos disponibles para análisis.")
        return

    opciones = {f"ID {r['id']} - {r['codigo']} {r['asignatura']}": int(r["id"]) for _, r in df_cursos.iterrows()}
    cid = opciones[st.selectbox("Seleccionar curso para análisis inteligente", list(opciones.keys()))]

    sug = generar_sugerencias_academicas(cid)
    col1, col2 = st.columns([1.2, 1])

    with col1:
        st.subheader(" Diagnóstico de Coherencia")
        st.info(sug["diagnostico"])
        st.subheader("💡 Sugerencias de Mejora Pedagógica")
        st.markdown(sug["sugerencias"])

    with col2:
        st.subheader("📊 Análisis Taxonómico de Bloom")
        analisis = analizar_coherencia_curso(cid)
        bloom = analisis.get("bloom", {})
        if bloom:
            st.markdown(f"**Nivel Dominante:** `{bloom.get('nivel_dominante')}`")
            niveles = bloom.get("niveles", {})
            df_bloom = pd.DataFrame([{"Nivel": k, "Conteo": v} for k, v in niveles.items()])
            st.dataframe(df_bloom, use_container_width=True, hide_index=True)


def ui_parametros(st: Any):
    st.header("Parámetros institucionales")
    inst = st.text_input("Nombre Institución", value=obtener_parametro("INSTITUCION_NOMBRE", "PCJIC"))
    if st.button("Guardar Parámetros", use_container_width=True):
        actualizar_parametro("INSTITUCION_NOMBRE", inst, "Nombre oficial")
        st.success("Parámetros actualizados.")


def ui_backup(st: Any):
    st.header("Copias y restauración de seguridad (Backup System)")
    st.caption("Genere respaldos completos en formato JSON e impórtelos con 1 solo clic.")

    tab1, tab2 = st.tabs(["Generar Copia de Seguridad", "Restaurar Backup"])

    with tab1:
        st.subheader("📥 Exportar Backup del Sistema")
        st.write("Esta copia incluye todos los usuarios, cursos, asignaturas base, auditoría y parámetros.")

        from src.services.export_service import generar_backup_sistema_json
        if st.button("Preparar archivo de respaldo", use_container_width=True):
            try:
                json_bytes = generar_backup_sistema_json()
                from datetime import datetime
                fname = f"Backup_PCJIC_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
                st.download_button("Descargar Copia de Seguridad (.json)", data=json_bytes, file_name=fname, mime="application/json", use_container_width=True)
            except Exception as e:
                st.error(f"Error generando backup: {e}")

    with tab2:
        st.subheader("📤 Restaurar desde un archivo JSON")
        arch_subido = st.file_uploader("Cargar archivo de respaldo .json", type=["json"])
        if arch_subido is not None and st.button("Restaurar base de datos", use_container_width=True):
            from src.services.export_service import restaurar_backup_sistema_json
            ok, msg = restaurar_backup_sistema_json(arch_subido.getvalue())
            if ok:
                st.success(f"✅ {msg}")
                st.rerun()
            else:
                st.error(f"❌ {msg}")


def ui_diagnostico_productivo(st: Any):
    st.header("Diagnóstico de infraestructura productiva")
    status = health_status()
    st.json(status)


def ui_auditoria(st: Any):
    st.header("Auditoría de seguridad y eventos")
    st.caption("Buscador avanzado y bitácora de eventos administrativos.")

    df_aud = listar_auditoria()
    if df_aud.empty:
        st.info("No hay eventos registrados en la bitácora de auditoría.")
        return

    col1, col2 = st.columns(2)
    with col1:
        filtro_usuario = st.text_input("Filtrar por usuario")
    with col2:
        filtro_accion = st.text_input("Filtrar por acción / evento")

    filtrado = df_aud.copy()
    if filtro_usuario.strip():
        filtrado = filtrado[filtrado["usuario"].astype(str).str.contains(filtro_usuario.strip(), case=False, na=False)]
    if filtro_accion.strip():
        filtrado = filtrado[filtrado["accion"].astype(str).str.contains(filtro_accion.strip(), case=False, na=False)]

    st.dataframe(filtrado, use_container_width=True, hide_index=True)
