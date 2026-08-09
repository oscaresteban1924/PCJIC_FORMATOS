import pytest
from datetime import date
from src.services.academic_service import (
    calcular_pascua,
    obtener_festivos_colombia,
    generar_fechas_clase,
)
import pandas as pd


def test_calcular_pascua():
    p2026 = calcular_pascua(2026)
    assert p2026 == date(2026, 4, 5)


def test_obtener_festivos_colombia_2026():
    festivos = obtener_festivos_colombia([2026])
    assert date(2026, 1, 1) in festivos   # Año nuevo
    assert date(2026, 4, 2) in festivos   # Jueves Santo
    assert date(2026, 4, 3) in festivos   # Viernes Santo
    assert date(2026, 7, 20) in festivos  # Independencia
    assert date(2026, 8, 7) in festivos   # Boyacá
    assert date(2026, 12, 25) in festivos # Navidad


def test_generar_fechas_clase_excluyendo_festivos():
    inicio = date(2026, 7, 15)
    fin = date(2026, 8, 15)
    horarios = pd.DataFrame([{"Día": "Viernes", "Hora inicio": "08:00", "Hora fin": "10:00", "Lugar / ambiente": "Aula 1"}])
    fechas_con_festivos = generar_fechas_clase(inicio, fin, horarios, incluir_festivos_colombia=True)

    fechas_list = fechas_con_festivos["Fecha"].tolist()
    assert date(2026, 8, 7) not in fechas_list  # Batalla de Boyacá es viernes 7 de Agosto 2026
