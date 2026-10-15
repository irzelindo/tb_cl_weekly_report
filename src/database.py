import os
from pathlib import Path
from typing import Any

import pyodbc
from dotenv import load_dotenv

# ---------------------------------------------------------
# Caminho base do projecto
# ---------------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent.parent


# ---------------------------------------------------------
# Carregar variáveis do .env
# ---------------------------------------------------------

load_dotenv(BASE_DIR / ".env")


def get_connection() -> pyodbc.Connection:
    """
    Cria e devolve uma ligação ao SQL Server.
    """

    driver = os.getenv("DB_DRIVER", "ODBC Driver 18 for SQL Server")

    server = os.getenv("DB_SERVER")
    database = os.getenv("DB_NAME", "TBData")
    username = os.getenv("DB_USER")
    password = os.getenv("DB_PASSWORD")

    encrypt = os.getenv("DB_ENCRYPT", "no")
    trust_certificate = os.getenv("DB_TRUST_SERVER_CERTIFICATE", "yes")

    if not server:
        raise ValueError("DB_SERVER não está definido no ficheiro .env")

    if not username:
        raise ValueError("DB_USER não está definido no ficheiro .env")

    if not password:
        raise ValueError("DB_PASSWORD não está definido no ficheiro .env")

    connection_string = (
        f"DRIVER={driver};"
        f"SERVER={server};"
        f"DATABASE={database};"
        f"UID={username};"
        f"PWD={password};"
        f"Encrypt={encrypt};"
        f"TrustServerCertificate={trust_certificate};"
        f"Connection Timeout=30;"
    )

    # print(connection_string)

    return pyodbc.connect(connection_string, autocommit=False)


def execute_query(sql: str, params: tuple | None = None) -> list[dict]:

    params = params or ()

    connection = get_connection()

    try:
        cursor = connection.cursor()

        cursor.execute(sql, params)

        if cursor.description is None:
            return []

        columns = [column[0] for column in cursor.description]

        results = []

        for row in cursor.fetchall():
            results.append(dict(zip(columns, row)))

        return results

    finally:
        connection.close()


def execute_query_file(
    sql_file: str | Path, params: tuple[Any, ...] | None = None
) -> list[dict]:
    """
    Lê uma query de um ficheiro .sql e executa-a.
    """

    sql_path = Path(sql_file)

    if not sql_path.is_absolute():
        sql_path = BASE_DIR / sql_path

    if not sql_path.exists():
        raise FileNotFoundError(f"Ficheiro SQL não encontrado: {sql_path}")

    sql = sql_path.read_text(encoding="utf-8-sig")

    return execute_query(sql=sql, params=params)
