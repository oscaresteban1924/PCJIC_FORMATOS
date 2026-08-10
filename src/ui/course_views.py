from __future__ import annotations

import io
import re
from datetime import date, datetime, timedelta
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

from src.config import (
    COLUMNAS_EVALUACIONES,
    COLUMNAS_GC72,
    COLUMNAS_HORARIOS,
    COLUMNAS_MODULOS,
    COLUMNAS_SESIONES,
    PREFORMAS_GC72,
    TEXTOS_PREDEFINIDOS_GC71,
)
from src.repositories.audit_repository import registrar_auditoria
from src.repositories.course_repository import (
    cambiar_estado_curso,
    df_to_payload,
    eliminar_curso,
    get_curso,
    guardar_version_curso,
    listar_cursos_visibles,
    observaciones_curso,
    payload_to_df,
    safe_json_loads,
    upsert_curso,
    versiones_curso,
)
from src.repositories.evidence_repository import evidencias_count, listar_evidencias, registrar_artefacto
from src.services.academic_service import (
    analizar_coherencia_curso,
    expandir_plan_sesiones,
    generar_fechas_clase,
    generar_sugerencias_academicas,
    limpiar_df,
    limpiar_numero,
    matriz_riesgo_cursos,
    score_calidad_expediente,
    validar_plan,
)
from src.services.calendar_service import build_ics_calendar
from src.services.docx_service import (
    construir_analisis,
    crear_gc71_docx,
    crear_informe_gc72_docx,
    normalizar_dataframe_gc72,
    texto_preformas,
)
from src.services.excel_service import (
    crear_plantilla_evaluacion_xlsx,
    diagnostico_tabla_estudiantes,
    leer_calificaciones,
    leer_excel_inteligente,
    leer_listado_estudiantes,
    normalizar_estudiantes_inteligente,
)
from src.services.export_service import crear_paquete_curso_zip, nombre_archivo_seguro, safe_filename


def df_vacio(columnas: List[str], filas: int = 5) -> pd.DataFrame:
    data = [{col: "" for col in columnas} for _ in range(filas)]
    return pd.DataFrame(data, columns=columnas)


def curso_key(row: Any, idx: int) -> str:
    codigo = re.sub(r"\W+", "_", str(row.get("Código", "")).strip())
    grupo = re.sub(r"\W+", "_", str(row.get("Grupo", "")).strip())
    asignatura = re.sub(r"\W+", "_", str(row.get("Asignatura", "")).strip())[:40]
    return f"curso_{idx}_{codigo}_{grupo}_{asignatura}"


from src.repositories.subject_repository import listar_asignaturas_base


def ui_gc71(st: Any):
    with st.expander("Qué hace este módulo", expanded=False):
        st.markdown("Genera la guía didáctica FD-GC71, cronograma automático según horario, evaluación concertada y plantilla Excel de evaluación.")

    st.header("FD-GC71 - Guía didáctica y concertación de evaluación")
    usuario_actual = st.session_state.get("auth_user", {})

    # Selector para cargar programa / microcurrículo del Banco Institucional o Expediente
    df_banco = listar_asignaturas_base()
    df_exp = listar_cursos_visibles(usuario_actual)

    col_b, col_e = st.columns([1.2, 1])
    sel_b_val = None
    sel_e_val = None

    with col_b:
        if not df_banco.empty:
            opciones_banco = {"-- Seleccionar del Banco Institucional --": None}
            for _, r_b in df_banco.iterrows():
                lbl = f"🏛️ {r_b.get('codigo','')} - {r_b.get('nombre','')} | {r_b.get('programa','')}"
                opciones_banco[lbl] = int(r_b["id"]) if "id" in r_b and r_b["id"] != "" else r_b.to_dict()
            sel_b_name = st.selectbox("🏛️ Seleccionar microcurrículo / Programa del Banco", list(opciones_banco.keys()), key="sel_banco_gc71")
            sel_b_val = opciones_banco[sel_b_name]
        else:
            st.caption("💡 No hay asignaturas en el Banco Institucional.")

    with col_e:
        if not df_exp.empty:
            opciones_exp = {"-- Cargar del Expediente Académico --": None}
            for _, r_e in df_exp.iterrows():
                lbl = f"📁 ID {r_e.get('id')} - {r_e.get('codigo','')} {r_e.get('asignatura','')} ({r_e.get('periodo','')})"
                opciones_exp[lbl] = int(r_e["id"])
            sel_e_name = st.selectbox("📁 O seleccionar curso guardado", list(opciones_exp.keys()), key="sel_exp_gc71")
            sel_e_val = opciones_exp[sel_e_name]

    # Determinar si la selección cambió para actualizar session_state
    current_key = f"exp_{sel_e_val}" if sel_e_val else (f"banco_{sel_b_val}" if sel_b_val else "blank")
    last_key = st.session_state.get("gc71_last_key", None)

    if last_key != current_key:
        st.session_state["gc71_last_key"] = current_key

        d = {}
        payload = {}
        if sel_e_val:
            curso = get_curso(sel_e_val)
            if curso:
                payload = safe_json_loads(curso.get("payload_json"), {})
                d = payload.get("datos", {})
                if not d:
                    d = {
                        "codigo": curso.get("codigo", ""),
                        "grupo": curso.get("grupo", ""),
                        "asignatura": curso.get("asignatura", ""),
                        "programa": curso.get("programa", ""),
                        "periodo": curso.get("periodo", ""),
                        "profesor": curso.get("profesor", ""),
                        "correo": curso.get("email_profesor", ""),
                        "creditos": curso.get("creditos", 3),
                        "htp": curso.get("htp", 2.0),
                        "hti": curso.get("hti", 4.0),
                    }
        elif sel_b_val:
            if isinstance(sel_b_val, dict):
                d = sel_b_val
            else:
                from src.repositories.subject_repository import get_asignatura_base
                d = get_asignatura_base(sel_b_val) or {}

        if d:
            st.session_state["input_programa"] = d.get("programa", "")
            st.session_state["input_asignatura"] = d.get("asignatura") or d.get("nombre") or ""
            st.session_state["input_codigo"] = d.get("codigo", "")
            st.session_state["input_area"] = d.get("area") or d.get("area_formacion") or ""
            st.session_state["input_profesor"] = d.get("profesor") or usuario_actual.get("nombre_completo", "")
            st.session_state["input_cedula"] = d.get("cedula_docente", "")
            st.session_state["input_correo"] = d.get("correo") or d.get("email_profesor") or usuario_actual.get("email", "")
            st.session_state["input_grupo"] = d.get("grupo", "")
            st.session_state["input_periodo"] = d.get("periodo", "2026-1")

            tipo_val = d.get("tipo_asignatura", "Teórico-práctica")
            st.session_state["input_tipo"] = tipo_val if tipo_val in ["Teórica", "Teórico-práctica", "Práctica"] else "Teórico-práctica"

            try:
                st.session_state["input_creditos"] = int(float(d.get("creditos", 3)))
            except Exception:
                st.session_state["input_creditos"] = 3

            try:
                st.session_state["input_htp"] = float(d.get("htp", 2.0))
                st.session_state["input_hti"] = float(d.get("hti", 4.0))
            except Exception:
                st.session_state["input_htp"] = 2.0
                st.session_state["input_hti"] = 4.0

            st.session_state["input_prerrequisitos"] = d.get("prerrequisitos", "Ninguno")
            st.session_state["input_correquisitos"] = d.get("correquisitos", "Ninguno")

            st.session_state["input_justificacion"] = d.get("justificacion") or TEXTOS_PREDEFINIDOS_GC71["justificacion"]
            st.session_state["input_competencias"] = d.get("competencias") or TEXTOS_PREDEFINIDOS_GC71["competencias"]
            st.session_state["input_resultados"] = d.get("resultados") or TEXTOS_PREDEFINIDOS_GC71["resultados"]
            st.session_state["input_objetivo_general"] = d.get("objetivo_general") or d.get("objetivos") or TEXTOS_PREDEFINIDOS_GC71["objetivo_general"]
            st.session_state["input_objetivos_especificos"] = d.get("objetivos_especificos") or TEXTOS_PREDEFINIDOS_GC71["objetivos_especificos"]
            st.session_state["input_metodologias"] = d.get("metodologias") or d.get("metodologia") or TEXTOS_PREDEFINIDOS_GC71["metodologias"]
            st.session_state["input_ambientes"] = d.get("ambientes") or TEXTOS_PREDEFINIDOS_GC71["ambientes"]
            st.session_state["input_medios"] = d.get("medios") or TEXTOS_PREDEFINIDOS_GC71["medios"]
            st.session_state["input_referencias"] = d.get("referencias") or d.get("bibliografia") or TEXTOS_PREDEFINIDOS_GC71["referencias"]

            if payload.get("modulos"):
                st.session_state["loaded_modulos_df"] = payload_to_df(payload["modulos"], COLUMNAS_MODULOS)
            else:
                st.session_state["loaded_modulos_df"] = None

            if payload.get("evaluaciones"):
                st.session_state["loaded_evaluaciones_df"] = payload_to_df(payload["evaluaciones"], COLUMNAS_EVALUACIONES)
            else:
                st.session_state["loaded_evaluaciones_df"] = None

        st.rerun()

    # Inicializar valores por defecto en session_state si es la primera vez
    if "input_programa" not in st.session_state:
        st.session_state["input_programa"] = ""
        st.session_state["input_asignatura"] = ""
        st.session_state["input_codigo"] = ""
        st.session_state["input_area"] = ""
        st.session_state["input_profesor"] = usuario_actual.get("nombre_completo", "")
        st.session_state["input_cedula"] = ""
        st.session_state["input_correo"] = usuario_actual.get("email", "")
        st.session_state["input_grupo"] = ""
        st.session_state["input_periodo"] = "2026-1"
        st.session_state["input_tipo"] = "Teórico-práctica"
        st.session_state["input_creditos"] = 3
        st.session_state["input_htp"] = 2.0
        st.session_state["input_hti"] = 4.0
        st.session_state["input_prerrequisitos"] = "Ninguno"
        st.session_state["input_correquisitos"] = "Ninguno"

        st.session_state["input_justificacion"] = TEXTOS_PREDEFINIDOS_GC71["justificacion"]
        st.session_state["input_competencias"] = TEXTOS_PREDEFINIDOS_GC71["competencias"]
        st.session_state["input_resultados"] = TEXTOS_PREDEFINIDOS_GC71["resultados"]
        st.session_state["input_objetivo_general"] = TEXTOS_PREDEFINIDOS_GC71["objetivo_general"]
        st.session_state["input_objetivos_especificos"] = TEXTOS_PREDEFINIDOS_GC71["objetivos_especificos"]
        st.session_state["input_metodologias"] = TEXTOS_PREDEFINIDOS_GC71["metodologias"]
        st.session_state["input_ambientes"] = TEXTOS_PREDEFINIDOS_GC71["ambientes"]
        st.session_state["input_medios"] = TEXTOS_PREDEFINIDOS_GC71["medios"]
        st.session_state["input_referencias"] = TEXTOS_PREDEFINIDOS_GC71["referencias"]

    with st.expander("1. Identificación de la asignatura", expanded=True):
        c1, c2, c3 = st.columns([1.2, 1.2, 0.8])
        with c1:
            programa = st.text_input("Programa académico", key="input_programa")
            asignatura = st.text_input("Asignatura", key="input_asignatura")
            codigo = st.text_input("Código", key="input_codigo")
            area = st.text_input("Área de formación", key="input_area")
        with c2:
            profesor = st.text_input("Profesor", key="input_profesor")
            cedula_docente = st.text_input("Cédula docente", key="input_cedula")
            correo = st.text_input("Correo electrónico", key="input_correo")
            grupo = st.text_input("Grupo", key="input_grupo")
        with c3:
            periodo = st.text_input("Periodo académico", key="input_periodo")
            tipo_opts = ["Teórica", "Teórico-práctica", "Práctica"]
            tipo_curr = st.session_state.get("input_tipo", "Teórico-práctica")
            tipo_idx = tipo_opts.index(tipo_curr) if tipo_curr in tipo_opts else 1
            tipo = st.selectbox("Tipo de asignatura", tipo_opts, index=tipo_idx, key="input_tipo")
            creditos = st.number_input("Número de créditos", min_value=0, step=1, key="input_creditos")
            htp = st.number_input("HTP semanal", min_value=0.0, step=0.5, key="input_htp")
            hti = st.number_input("HTI semanal", min_value=0.0, step=0.5, key="input_hti")
        prerrequisitos = st.text_input("Prerrequisito(s)", key="input_prerrequisitos")
        correquisitos = st.text_input("Correquisito(s)", key="input_correquisitos")

    with st.expander("2. Textos académicos base", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            justificacion = st.text_area("Justificación", key="input_justificacion", height=140)
            competencias = st.text_area("Competencias a las que tributa", key="input_competencias", height=130)
            resultados = st.text_area("Resultados de aprendizaje", key="input_resultados", height=130)
            objetivo_general = st.text_area("Objetivo general", key="input_objetivo_general", height=110)
        with col2:
            objetivos_especificos = st.text_area("Objetivos específicos", key="input_objetivos_especificos", height=140)
            metodologias = st.text_area("Metodologías y estrategias didácticas", key="input_metodologias", height=130)
            ambientes = st.text_area("Ambientes de aprendizaje", key="input_ambientes", height=100)
            medios = st.text_area("Medios educativos", key="input_medios", height=100)
            referencias = st.text_area("Referencias bibliográficas", key="input_referencias", height=100)

    st.subheader("3. Módulos / unidades e intensidad")
    criterio = st.radio("Criterio de expansión del cronograma", ["Horas presenciales", "Sesiones"], horizontal=True)
    default_modulos = pd.DataFrame([
        {"Unidad": "UNIDAD 1. Fundamentos", "Contenido / tema central": "Introducción, conceptos base y alcance de la asignatura", "Horas presenciales": 4, "Sesiones": 2, "Trabajo presencial": "Clase orientadora, explicación conceptual y taller diagnóstico.", "Trabajo independiente": "Lectura de apoyo y preparación de preguntas orientadoras."},
        {"Unidad": "UNIDAD 2. Aplicación", "Contenido / tema central": "Desarrollo de procedimientos, ejercicios y análisis de casos", "Horas presenciales": 8, "Sesiones": 4, "Trabajo presencial": "Taller aplicado, solución de ejercicios y discusión guiada.", "Trabajo independiente": "Desarrollo de actividad práctica y revisión bibliográfica."},
        {"Unidad": "UNIDAD 3. Integración", "Contenido / tema central": "Proyecto, socialización y retroalimentación", "Horas presenciales": 4, "Sesiones": 2, "Trabajo presencial": "Acompañamiento al proyecto y socialización de resultados.", "Trabajo independiente": "Ajuste de entregables y preparación de sustentación."},
    ])

    modulos_initial = st.session_state.get("loaded_modulos_df")
    if modulos_initial is None or (isinstance(modulos_initial, pd.DataFrame) and modulos_initial.empty):
        modulos_initial = default_modulos

    modulos_df = st.data_editor(modulos_initial, num_rows="dynamic", hide_index=True, use_container_width=True, key="modulos_editor_gc71")

    st.subheader("4. Horario de clase")
    c1, c2, c3 = st.columns(3)
    with c1:
        fecha_inicio = st.date_input("Fecha de inicio", value=date.today(), format="DD/MM/YYYY", key="gc71_inicio")
    with c2:
        fecha_fin = st.date_input("Fecha de finalización", value=date.today() + timedelta(days=112), format="DD/MM/YYYY", key="gc71_fin")
    with c3:
        fechas_no_clase_txt = st.text_area("Fechas sin clase (dd/mm/aaaa, una por línea)", value="", height=100)

    excluir_festivos_co = st.checkbox("Excluir festivos nacionales de Colombia automáticamente (Ley Emiliani y Fiestas Móviles)", value=True, key="gc71_festivos_co")

    default_horarios = pd.DataFrame([
        {"Día": "Lunes", "Hora inicio": "18:00", "Hora fin": "20:00", "Lugar / ambiente": "Aula de clase"},
    ])
    horarios_df = st.data_editor(default_horarios, num_rows="dynamic", hide_index=True, use_container_width=True, key="horarios_gc71")

    fechas_excluidas = []
    for line in (fechas_no_clase_txt or "").split("\n"):
        txt = line.strip()
        if txt:
            try:
                fechas_excluidas.append(datetime.strptime(txt, "%d/%m/%Y").date())
            except Exception:
                pass

    fechas_clase_df = generar_fechas_clase(fecha_inicio, fecha_fin, horarios_df, fechas_excluidas, incluir_festivos_colombia=excluir_festivos_co)
    sesiones_df = expandir_plan_sesiones(modulos_df, fechas_clase_df, criterio=criterio)

    st.subheader("5. Plan de sesiones (generado automáticamente)")
    st.write(f"Sesiones calculadas: **{len(sesiones_df)}** | Fechas válidas de clase: **{len(fechas_clase_df)}**")
    sesiones_df = st.data_editor(sesiones_df, num_rows="dynamic", hide_index=True, use_container_width=True, key="sesiones_editor_gc71")

    st.subheader("6. Concertación de evaluación")
    default_eval = pd.DataFrame([
        {"Tipo de evaluación": "Primer parcial", "Procedimiento de evaluación": "Evaluación escrita individual", "Valor (%)": 25, "Fecha de realización": "15/09/2026", "Unidad relacionada": "UNIDAD 1", "Corte": "Primer corte"},
        {"Tipo de evaluación": "Seguimiento y talleres", "Procedimiento de evaluación": "Talleres individuales y grupales", "Valor (%)": 25, "Fecha de realización": "10/10/2026", "Unidad relacionada": "UNIDAD 2", "Corte": "Segundo corte"},
        {"Tipo de evaluación": "Segundo parcial", "Procedimiento de evaluación": "Evaluación teórica y práctica", "Valor (%)": 25, "Fecha de realización": "01/11/2026", "Unidad relacionada": "UNIDAD 2", "Corte": "Segundo corte"},
        {"Tipo de evaluación": "Proyecto final", "Procedimiento de evaluación": "Entrega final y sustentación", "Valor (%)": 25, "Fecha de realización": "20/11/2026", "Unidad relacionada": "UNIDAD 3", "Corte": "Tercer corte"},
    ])

    eval_initial = st.session_state.get("loaded_evaluaciones_df")
    if eval_initial is None or (isinstance(eval_initial, pd.DataFrame) and eval_initial.empty):
        eval_initial = default_eval

    evaluaciones_df = st.data_editor(eval_initial, num_rows="dynamic", hide_index=True, use_container_width=True, key="evaluaciones_editor_gc71")

    datos = {
        "programa": programa,
        "asignatura": asignatura,
        "codigo": codigo,
        "area": area,
        "profesor": profesor,
        "cedula_docente": cedula_docente,
        "correo": correo,
        "grupo": grupo,
        "periodo": periodo,
        "tipo_asignatura": tipo,
        "creditos": str(creditos),
        "htp": str(htp),
        "hti": str(hti),
        "ht_total": str(htp + hti),
        "prerrequisitos": prerrequisitos,
        "correquisitos": correquisitos,
        "justificacion": justificacion,
        "competencias": competencias,
        "resultados": resultados,
        "objetivo_general": objetivo_general,
        "objetivos_especificos": objetivos_especificos,
        "metodologias": metodologias,
        "ambientes": ambientes,
        "medios": medios,
        "referencias": referencias,
        "fecha_socializacion": date.today().strftime("%d/%m/%Y"),
        "fecha_revision": date.today().strftime("%d/%m/%Y"),
        "fecha_aprobacion": date.today().strftime("%d/%m/%Y"),
    }

    st.subheader("7. Generar y guardar")
    alertas, recomendaciones = validar_plan(modulos_df, horarios_df, sesiones_df, evaluaciones_df, datos)
    for a in alertas:
        st.error(f"⚠️ {a}")
    for r in recomendaciones:
        st.info(f"💡 {r}")

    col_btn1, col_btn2, col_btn3, col_btn4 = st.columns(4)
    with col_btn1:
        if st.button("Guardar en Expediente Académico", use_container_width=True):
            try:
                payload = {
                    "datos": datos,
                    "sesiones": df_to_payload(sesiones_df),
                    "evaluaciones": df_to_payload(evaluaciones_df),
                    "modulos": df_to_payload(modulos_df),
                }
                cid = upsert_curso(None, datos, payload, user=usuario_actual)
                registrar_auditoria("Guardar FD-GC71", f"Curso ID={cid} - {datos.get('asignatura')}", user=usuario_actual)
                st.success(f"Curso guardado en expediente con ID {cid}.")
            except Exception as e:
                st.error(f"Error al guardar: {e}")

    with col_btn2:
        try:
            docx_bytes = crear_gc71_docx(datos, sesiones_df, evaluaciones_df)
            st.download_button("Descargar FD-GC71 Word (.docx)", data=docx_bytes, file_name=f"FD-GC71_{codigo or 'Curso'}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
        except Exception as e:
            st.error(f"Error generando Word: {e}")

    with col_btn3:
        try:
            from src.services.docx_service import crear_gc71_html_imprimible
            html_content = crear_gc71_html_imprimible(datos, sesiones_df, evaluaciones_df)
            st.download_button("🖨️ Imprimir / PDF (HTML)", data=html_content, file_name=f"FD-GC71_{codigo or 'Curso'}.html", mime="text/html", use_container_width=True)
        except Exception as e:
            st.error(f"Error generando versión HTML: {e}")

    with col_btn4:
        try:
            from src.services.calendar_service import build_ics_reminders
            rem_bytes = build_ics_reminders(evaluaciones_df, datos)
            st.download_button("⏰ Recordatorios iCal (.ics)", data=rem_bytes, file_name=f"Recordatorios_{codigo or 'Curso'}.ics", mime="text/calendar", use_container_width=True)
        except Exception as e:
            st.error(f"Error generando recordatorios: {e}")


def ui_gc72(st: Any):
    with st.expander("Qué hace este módulo", expanded=False):
        st.markdown("Genera el informe académico FD-GC72 usando listado tradicional, notas de corte, análisis descriptivo e indicadores estadísticos.")

    st.header("FD-GC72 - Informe académico")
    with st.sidebar:
        st.markdown("### Opciones FD-GC72")
        calcular = st.checkbox("Calcular porcentajes automáticamente", value=True)
        modo_analisis = st.radio("Tipo de análisis descriptivo", ["Por cada curso", "Consolidado institucional"], index=0)

    st.subheader("1. Datos generales")
    c1, c2, c3 = st.columns([1.4, 1, 1])
    with c1:
        docente = st.text_input("Docente", value="", key="gc72_docente")
    with c2:
        periodo = st.text_input("Período académico", value="2026-1", key="gc72_periodo")
    with c3:
        fecha_entrega = st.date_input("Fecha de entrega", value=date.today(), format="DD/MM/YYYY", key="gc72_fecha")

    st.subheader("2. Cursos reportados e Indicadores Estadísticos")
    base = df_vacio(COLUMNAS_GC72, filas=3)
    cursos_editados = st.data_editor(base, hide_index=True, num_rows="dynamic", use_container_width=True, key="tabla_gc72")
    cursos = normalizar_dataframe_gc72(cursos_editados, calcular_porcentajes=calcular)

    if not cursos.empty:
        from src.services.excel_service import calcular_estadisticas_curso
        stats = calcular_estadisticas_curso(cursos, columna_nota="Aprueban a la fecha N°")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Estudiantes Matriculados", stats.get("total", 0))
        m2.metric("Media de Aprobación", f"{stats.get('media', 0)}")
        m3.metric("Desviación Estándar", stats.get("desviacion", 0))
        m4.metric("Tasa de Aprobación Global", stats.get("tasa_aprobacion", "0%"))

    if cursos.empty:
        st.info("Agregue al menos un curso para activar el análisis y la descarga.")
        return

    st.subheader("3. Análisis descriptivo")
    analisis_por_curso: Dict[str, Dict[str, str]] = {}
    for idx, row in cursos.iterrows():
        key = curso_key(row, idx)
        titulo = str(row.get("Asignatura", "")).strip() or f"Curso {idx + 1}"
        with st.expander(f"{titulo} (Código: {row.get('Código','')})", expanded=(idx == 0)):
            c_a, c_b, c_c = st.columns(3)
            with c_a:
                sel_pos = st.multiselect("Aspectos positivos", list(PREFORMAS_GC72["aspectos_positivos"].keys()), default=["Avance adecuado", "Aplicación práctica"], key=f"pos_{key}")
                add_pos = st.text_area("Adicional positivo", key=f"pos_add_{key}", height=80)
            with c_b:
                sel_inc = st.multiselect("Inconvenientes", list(PREFORMAS_GC72["inconvenientes"].keys()), default=[], key=f"inc_{key}")
                add_inc = st.text_area("Adicional inconvenientes", key=f"inc_add_{key}", height=80)
            with c_c:
                sel_prop = st.multiselect("Propuestas", list(PREFORMAS_GC72["propuestas"].keys()), default=["Seguimiento formativo"], key=f"prop_{key}")
                add_prop = st.text_area("Adicional propuestas", key=f"prop_add_{key}", height=80)
            analisis_por_curso[key] = {
                "positivos": texto_preformas(sel_pos, "aspectos_positivos", add_pos),
                "inconvenientes": texto_preformas(sel_inc, "inconvenientes", add_inc),
                "propuestas": texto_preformas(sel_prop, "propuestas", add_prop),
            }

    bloques = construir_analisis(cursos, analisis_por_curso, modo_analisis)

    st.subheader("4. Descargar")
    if not docente.strip():
        st.warning("Ingrese el nombre del docente antes de descargar.")
        return
    try:
        docx_bytes = crear_informe_gc72_docx(docente, periodo, fecha_entrega, cursos, bloques)
        nombre_base = nombre_archivo_seguro(docente, fecha_entrega, "FD_GC72_Informe_Academico")
        st.download_button("Descargar FD-GC72 Word (.docx)", data=docx_bytes, file_name=f"{nombre_base}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
    except Exception as e:
        st.error(f"Error generando FD-GC72: {e}")


def ui_expediente_academico(st: Any):
    st.header("Expediente académico")
    user = st.session_state.get("auth_user", {})
    df_cursos = listar_cursos_visibles(user)
    if df_cursos.empty:
        st.info("No hay cursos registrados en el expediente.")
        return

    st.dataframe(df_cursos[["id", "codigo", "grupo", "asignatura", "programa", "periodo", "profesor", "estado", "actualizado_en"]], use_container_width=True, hide_index=True)

    opciones = {f"ID {r['id']} - {r['codigo']} {r['asignatura']} ({r['grupo']})": int(r["id"]) for _, r in df_cursos.iterrows()}
    sel = st.selectbox("Seleccionar expediente para ver detalle", list(opciones.keys()))
    cid = opciones[sel]

    c = get_curso(cid)
    if c:
        st.subheader(f"Detalle del expediente ID {cid}")
        st.json(safe_json_loads(c.get("payload_json")))
        score, diag = score_calidad_expediente(cid)
        st.metric("Puntuación de Calidad Documental", f"{score}/100")
        st.table(pd.DataFrame(diag, columns=["Criterio", "Estado", "Puntos"]))


def ui_planeador_superior(st: Any):
    st.header("Planeador superior de cursos")
    st.info("Utilice este módulo para armar cursos, expandir planes de clases, clonar periodos anteriores y exportar el paquete integrado en ZIP.")

    with st.expander("🔄 Clonar curso desde periodo anterior", expanded=False):
        user = st.session_state.get("auth_user", {})
        df_c = listar_cursos_visibles(user)
        if not df_c.empty:
            c_dict = {f"ID {r['id']} - {r['codigo']} {r['asignatura']} ({r['periodo']})": int(r["id"]) for _, r in df_c.iterrows()}
            src_id = c_dict[st.selectbox("Seleccione curso a duplicar", list(c_dict.keys()), key="clon_src")]
            target_period = st.text_input("Nuevo periodo de destino", value="2026-2", key="clon_target")
            if st.button("Clonar Syllabus y Planeación a Nuevo Periodo", use_container_width=True):
                try:
                    from src.repositories.course_repository import clonar_curso_a_nuevo_periodo
                    new_id = clonar_curso_a_nuevo_periodo(src_id, target_period, user=user)
                    st.success(f"✅ Curso clonado exitosamente para el periodo {target_period} con nuevo ID {new_id}.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al clonar curso: {e}")

    ui_gc71(st)


def ui_evidencias(st: Any):
    st.header("Gestión de evidencias y soportes firmados")
    user = st.session_state.get("auth_user", {})
    df_cursos = listar_cursos_visibles(user)
    if df_cursos.empty:
        st.info("No hay cursos disponibles.")
        return
    opciones = {f"ID {r['id']} - {r['codigo']} {r['asignatura']}": int(r["id"]) for _, r in df_cursos.iterrows()}
    cid = opciones[st.selectbox("Curso", list(opciones.keys()))]

    st.subheader("Cargar soporte o evidencia")
    uploaded = st.file_uploader("Archivo de evidencia (PDF, PNG, JPG, DOCX)", type=["pdf", "png", "jpg", "jpeg", "docx", "xlsx"])
    tipo = st.selectbox("Tipo de soporte", ["Guía firmada", "Concertación evaluativa", "Listado de asistencia", "Parciales", "Informe final", "Otro"])
    desc = st.text_input("Descripción opcional")
    if st.button("Guardar evidencia", use_container_width=True) and uploaded is not None:
        try:
            b64 = ""
            registrar_artefacto(cid, tipo, uploaded.name, uploaded.type, uploaded.size, b64, user=user)
            st.success("Evidencia guardada exitosamente.")
            st.rerun()
        except Exception as e:
            st.error(str(e))

    st.subheader("Evidencias del curso")
    df_ev = listar_evidencias(cid)
    st.dataframe(df_ev, use_container_width=True, hide_index=True)


def ui_cargador_inteligente(st: Any):
    st.header("Cargador inteligente de notas y listados")
    st.caption("Inspecciona automáticamente archivos Excel o CSV, identifica encabezados y valida duplicados.")
    uploaded = st.file_uploader("Cargar listado o plantilla Excel", type=["xlsx", "xls", "csv"])
    if uploaded is not None:
        tablas = leer_excel_inteligente(uploaded)
        for name, df in tablas.items():
            st.subheader(f"Hoja: {name}")
            st.dataframe(df.head(20), use_container_width=True)
            norm = normalizar_estudiantes_inteligente(df, {})

            from src.services.excel_service import detectar_duplicados_estudiantes
            dups = detectar_duplicados_estudiantes(norm)
            if not dups.empty:
                st.warning(f"⚠️ Se detectaron {len(dups)} registros duplicados por Documento o Nombre en el listado cargado:")
                st.dataframe(dups[["Documento", "Nombre completo", "Correo", "Estado"]], use_container_width=True, hide_index=True)

            diag = diagnostico_tabla_estudiantes(norm)
            st.dataframe(diag, use_container_width=True)


def ui_comparador_cortes(st: Any):
    st.header("Comparador de cortes académicos")
    st.info("Módulo de seguimiento comparativo de rendimiento por cortes evaluativos.")


def ui_semaforo_expediente(st: Any):
    st.header("Semáforo documental de riesgo")
    df_matriz = matriz_riesgo_cursos()
    if not df_matriz.empty:
        st.dataframe(df_matriz, use_container_width=True, hide_index=True)
    else:
        st.info("No hay datos de riesgo disponibles.")
