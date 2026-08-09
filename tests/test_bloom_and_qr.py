import pytest
import pandas as pd
from src.services.academic_service import analizar_taxonomia_bloom
from src.services.export_service import generar_codigo_qr_bytes, crear_paquete_curso_zip


def test_analizar_taxonomia_bloom():
    objetivos = "El estudiante logrará analizar, diseñar y proponer arquitecturas distribuidas."
    evaluaciones = pd.DataFrame([{"Tipo de evaluación": "Proyecto de diseño", "Procedimiento de evaluación": "Sustentación de caso"}])

    res = analizar_taxonomia_bloom(objetivos, evaluaciones)
    assert res["nivel_dominante"] in ("Analizar", "Crear")
    assert res["coherente"] is True


def test_generar_codigo_qr():
    qr_bytes = generar_codigo_qr_bytes("PRUEBA_INTEGRIDAD_PCJIC")
    assert qr_bytes is not None
    assert len(qr_bytes) > 0


def test_crear_paquete_zip_con_qr():
    datos = {"programa": "Software", "asignatura": "Pruebas", "codigo": "ISO1", "grupo": "01", "profesor": "Docente"}
    sesiones = pd.DataFrame([{"Unidad": "U1", "N° sesión": 1, "Fecha": "2026-08-10", "Horario": "08:00 - 10:00", "Contenido por desarrollar": "Tema 1", "Descripción del trabajo presencial": "TP", "Descripción trabajo independiente": "TI", "Lugar / ambiente": "Aula 1"}])
    evaluaciones = pd.DataFrame([{"Tipo de evaluación": "Parcial", "Procedimiento de evaluación": "Escrito", "Valor (%)": 100, "Fecha de realización": "2026-09-01", "Unidad relacionada": "U1", "Corte": "C1"}])
    estudiantes = pd.DataFrame([{"Nombre completo": "Est 1", "Documento": "123", "Correo": "c@c.com", "Plan": "P", "Observación": "", "Estado": "Activo"}])

    zip_bytes = crear_paquete_curso_zip(datos, sesiones, evaluaciones, estudiantes, curso_id=999)
    assert zip_bytes is not None
    assert len(zip_bytes) > 0
