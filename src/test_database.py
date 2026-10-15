from database import get_connection


def main():
    print("=" * 70)
    print("TESTE DE CONEXÃO AO SQL SERVER")
    print("=" * 70)

    try:
        connection = get_connection()
        cursor = connection.cursor()

        cursor.execute(
            """
            SELECT
                @@SERVERNAME AS ServerName,
                DB_NAME() AS DatabaseName,
                @@VERSION AS SQLServerVersion,
                GETDATE() AS ServerDateTime;
            """
        )

        row = cursor.fetchone()

        print()
        print("Conexão estabelecida com sucesso.")
        print()
        print(f"Servidor : {row.ServerName}")
        print(f"Database : {row.DatabaseName}")
        print(f"Data     : {row.ServerDateTime}")
        print()
        print("Versão SQL Server:")
        print(row.SQLServerVersion)

        connection.close()

        print()
        print("=" * 70)
        print("PASSOU")
        print("=" * 70)

    except Exception as exc:
        print()
        print("=" * 70)
        print("FALHOU")
        print("=" * 70)
        print()
        print(type(exc).__name__)
        print(exc)
        raise


if __name__ == "__main__":
    main()
