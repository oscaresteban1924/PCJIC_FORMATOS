import pytest
import pandas as pd
from datetime import date
from src.services.docx_service import crear_gc71_docx, crear_informe_gc72_docx
from src.config import TEXTOS_PREDEFINIDOS_GC71, COLUMNAS_GC72


def test_crear_gc71_docx():
    datos = {
        "programa": "Ingeniería de Software",
        "asignatura": "Pruebas de Software",
        "codigo": "ISO-501",
        "grupo": "01",
        "profesor": "Docente Test",
        "correo": "docente@test.edu.co",
        "periodo": "2026-1",
        "creditos": "3",
        "htp": 3,
        "hti": 6,
        **TEXTOS_PREDEFINIDOS_GC71,
    }
    sesiones = pd.DataFrame([
        {"Unidad": "U1", "N° sesión": 1, "Fecha": "2026-08-10", "Horario": "08:00 - 10:00", "Contenido por desarrollar": "Tema 1", "Descripción del trabajo presencial": "Clase", "Descripción trabajo independiente": "Lectura", "Lugar / ambiente": "Aula 101"}
    ])
    evaluaciones = pd.DataFrame([
        {"Tipo de evaluación": "Parcial", "Procedimiento de evaluación": "Escrito", "Valor (%)": 100, "Fecha de realización": "2026-10-10", "Unidad relacionada": "U1", "Corte": "Corte 1"}
    ])
    docx_bytes = crear_gc71_docx(datos, sesiones, evaluaciones)
    assert docx_bytes is not None
    assert len(docx_bytes) > 0


def test_crear_informe_gc72_docx():
    cursos = pd.DataFrame([{col: "10" if "%" not in col and "N°" not in col and col not in ["Código", "Grupo", "Asignatura"] else "10" for col in COLUMNAS_GC72}])
    cursos["Código"] = "ISO-501"
    cursos["Grupo"] = "01"
    cursos["Asignatura"] = "Pruebas"

    bloques = [{"titulo": "Curso 1", "positivos": "Buen avance", "inconvenientes": "Ninguno", "propuestas": "Seguir igual"}]
    docx_bytes = crear_informe_gc72_docx("Docente Test", "2026-1", date.today(), cursos, bloques)
    assert docx_bytes is not None
    assert len(docx_bytes) > 0
