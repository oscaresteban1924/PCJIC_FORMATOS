import pytest
from datetime import date
import pandas as pd
from src.services.academic_service import generar_fechas_clase, expandir_plan_sesiones, validar_plan


def test_generar_fechas_clase_y_expandir():
    inicio = date(2026, 8, 3)
    fin = date(2026, 8, 31)
    horarios = pd.DataFrame([{"Día": "Lunes", "Hora inicio": "08:00", "Hora fin": "10:00", "Lugar / ambiente": "Aula 1"}])
    fechas_clase = generar_fechas_clase(inicio, fin, horarios)

    assert not fechas_clase.empty
    assert len(fechas_clase) >= 4

    modulos = pd.DataFrame([{"Unidad": "U1", "Contenido / tema central": "Tema 1", "Horas presenciales": 4, "Sesiones": 2, "Trabajo presencial": "TP", "Trabajo independiente": "TI"}])
    sesiones = expandir_plan_sesiones(modulos, fechas_clase)
    assert not sesiones.empty
    assert len(sesiones) == 2


def test_validar_plan():
    modulos = pd.DataFrame([{"Unidad": "U1"}])
    horarios = pd.DataFrame([{"Día": "Lunes"}])
    sesiones = pd.DataFrame([{"Unidad": "U1"}])
    evaluaciones = pd.DataFrame([{"Valor (%)": 100}])
    datos = {"creditos": 3, "htp": 2, "hti": 4}

    alertas, recs = validar_plan(modulos, horarios, sesiones, evaluaciones, datos)
    assert len(alertas) == 0
