import pytest
import pandas as pd
from src.database.schema import init_db
from src.database.connection import conexion_db, db_execute, read_sql_df


def test_init_db():
    init_db()
    conn = conexion_db()
    assert conn is not None
    conn.close()


def test_db_execute_and_query():
    init_db()
    res = db_execute("SELECT COUNT(*) AS total FROM usuarios", fetchone=True)
    assert res is not None
    assert "total" in res or "TOTAL" in res or len(res) > 0

    df = read_sql_df("SELECT * FROM usuarios")
    assert isinstance(df, pd.DataFrame)
    assert not df.empty
