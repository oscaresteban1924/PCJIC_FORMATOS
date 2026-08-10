import pytest
import pandas as pd
from src.database.schema import init_db
from src.repositories.course_repository import upsert_curso, clonar_curso_a_nuevo_periodo
from src.services.export_service import generar_certificado_calidad_html
from src.services.excel_service import detectar_duplicados_estudiantes


def test_clonar_curso_a_nuevo_periodo():
    init_db()
    datos = {"codigo": "TEST1", "grupo": "01", "asignatura": "Algoritmos 1", "programa": "Sistemas", "periodo": "2025-2", "profesor": "Docente Test"}
    payload = {"datos": datos, "sesiones": [], "evaluaciones": []}
    cid = upsert_curso(None, datos, payload)

    new_id = clonar_curso_a_nuevo_periodo(cid, "2026-1")
    assert new_id is not None
    assert new_id != cid


def test_generar_certificado_calidad_html():
    init_db()
    datos = {"codigo": "CERT1", "grupo": "01", "asignatura": "Calidad de Software", "programa": "Sistemas", "periodo": "2026-1", "profesor": "Prof Cert"}
    payload = {"datos": datos, "sesiones": [], "evaluaciones": []}
    cid = upsert_curso(None, datos, payload)

    html = generar_certificado_calidad_html(cid)
    assert html is not None
    assert "Certificado de Calidad" in html
    assert "CALIDAD DE SOFTWARE" in html


def test_detectar_duplicados_estudiantes():
    df_est = pd.DataFrame([
        {"Documento": "101", "Documento llave": "101", "Nombre completo": "Juan Perez"},
        {"Documento": "102", "Documento llave": "102", "Nombre completo": "Maria Gomez"},
        {"Documento": "101", "Documento llave": "101", "Nombre completo": "Juan Perez Repetido"},
    ])

    dups = detectar_duplicados_estudiantes(df_est)
    assert not dups.empty
    assert len(dups) == 2
