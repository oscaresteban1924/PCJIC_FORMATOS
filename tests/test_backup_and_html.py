import pytest
import pandas as pd
from src.services.export_service import generar_backup_sistema_json, restaurar_backup_sistema_json
from src.services.docx_service import crear_gc71_html_imprimible


def test_generar_y_restaurar_backup():
    backup_bytes = generar_backup_sistema_json()
    assert backup_bytes is not None
    assert len(backup_bytes) > 0

    ok, msg = restaurar_backup_sistema_json(backup_bytes)
    assert ok is True
    assert "restaurada exitosamente" in msg


def test_crear_gc71_html_imprimible():
    datos = {"programa": "Sistemas", "asignatura": "Bases de Datos", "codigo": "BD1", "profesor": "Profesor Test"}
    sesiones = pd.DataFrame([{"Unidad": "U1", "N° sesión": 1, "Fecha": "2026-08-10", "Horario": "08:00 - 10:00", "Contenido por desarrollar": "Introducción"}])
    evaluaciones = pd.DataFrame([{"Tipo de evaluación": "Examen", "Procedimiento de evaluación": "Práctica", "Valor (%)": 100}])

    html = crear_gc71_html_imprimible(datos, sesiones, evaluaciones)
    assert html is not None
    assert "<html" in html
    assert "FD-GC71" in html
    assert "Bases de Datos" in html
