"""Perfil dos nomes reais das unidades executantes na dbo.TBCultura.

Não devolve identificadores pessoais.
"""
import argparse
from datetime import datetime, timedelta
from database import execute_query


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--week-start", required=True)
    args = parser.parse_args()
    begin = datetime.strptime(args.week_start, "%Y-%m-%d").date()
    end = begin + timedelta(days=7)
    sql = """
    SET NOCOUNT ON;
    DECLARE @WeekEnd DATE = ?;
    SELECT TOP (100)
        LTRIM(RTRIM(ISNULL(T.TestingFacilityName, ''))) AS TestingFacilityName,
        LTRIM(RTRIM(ISNULL(T.LIMSFacilityName, ''))) AS LIMSFacilityName,
        COUNT_BIG(*) AS Registos
    FROM dbo.TBCultura AS T
    WHERE T.RegisteredDateTime < @WeekEnd
    GROUP BY T.TestingFacilityName, T.LIMSFacilityName
    ORDER BY COUNT_BIG(*) DESC;
    """
    result = execute_query(sql, (end,))
    print(f"Unidades executantes encontradas até {end}:")
    for row in result:
        print(
            f"{row.get('TestingFacilityName')!r:42} "
            f" / LIMS={row.get('LIMSFacilityName')!r:42}"
            f" / n={row.get('Registos')}"
        )
    print("Verificar as designações de Carmelo, Nampula, Maputo e Beira.")
    print("Confirmar que não há laboratórios pretendidos excluídos por alias.")


if __name__ == "__main__":
    main()
