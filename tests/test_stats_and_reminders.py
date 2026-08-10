import pytest
import pandas as pd
from src.services.excel_service import calcular_estadisticas_curso
from src.services.calendar_service import build_ics_reminders


def test_calcular_estadisticas_curso():
    df_notas = pd.DataFrame([
        {"Nota final": 4.5},
        {"Nota final": 3.5},
        {"Nota final": 2.0},
        {"Nota final": 5.0},
    ])

    stats = calcular_estadisticas_curso(df_notas, columna_nota="Nota final")
    assert stats["total"] == 4
    assert stats["media"] == 3.75
    assert stats["minimo"] == 2.0
    assert stats["maximo"] == 5.0
    assert stats["aprueban"] == 3
    assert stats["reprueban"] == 1
    assert stats["tasa_aprobacion"] == "75.0%"


def test_build_ics_reminders():
    evaluaciones = pd.DataFrame([
        {"Tipo de evaluación": "Parcial 1", "Valor (%)": 25, "Fecha de realización": "2026-09-15", "Procedimiento de evaluación": "Escrito", "Corte": "Corte 1"}
    ])
    datos = {"asignatura": "Algoritmos", "codigo": "ALG1", "grupo": "01"}

    ics_bytes = build_ics_reminders(evaluaciones, datos)
    assert ics_bytes is not None
    assert b"VALARM" in ics_bytes
    assert b"TRIGGER:-P2D" in ics_bytes
