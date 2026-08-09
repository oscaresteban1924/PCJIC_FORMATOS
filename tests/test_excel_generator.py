import pytest
import pandas as pd
from src.services.excel_service import crear_plantilla_evaluacion_xlsx, normalizar_estudiantes_inteligente


def test_crear_plantilla_evaluacion_xlsx():
    datos = {"programa": "Prog", "asignatura": "Asig", "codigo": "COD", "grupo": "01", "periodo": "2026-1", "profesor": "Prof"}
    est = pd.DataFrame([{"Nombre completo": "Estudiante 1", "Documento": "123", "Correo": "est@test.com", "Plan": "P", "Observación": "", "Estado": "Activo"}])
    ev = pd.DataFrame([{"Tipo de evaluación": "Taller", "Procedimiento de evaluación": "Entregable", "Valor (%)": 100, "Fecha de realización": "2026-09-01", "Unidad relacionada": "U1", "Corte": "Corte 1"}])

    excel_bytes = crear_plantilla_evaluacion_xlsx(est, ev, datos)
    assert excel_bytes is not None
    assert len(excel_bytes) > 0


def test_normalizar_estudiantes_inteligente():
    df_raw = pd.DataFrame([
        {"CC / Documento": "1001", "Estudiante Nombre": "Juan Perez", "Correo": "juan@mail.com", "Nota Parcial": "4.5"},
    ])
    mapping = {"documento": "CC / Documento", "nombre": "Estudiante Nombre", "correo": "Correo", "nota_parcial": "Nota Parcial"}
    norm = normalizar_estudiantes_inteligente(df_raw, mapping)
    assert not norm.empty
    assert norm.iloc[0]["Documento"] == "1001"
    assert norm.iloc[0]["Nombre completo"] == "Juan Perez"
    assert norm.iloc[0]["Nota parcial"] == 4.5
