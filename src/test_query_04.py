from datetime import date
from pathlib import Path

from database import execute_query_file

BASE_DIR = Path(__file__).resolve().parent.parent


def main():
    week_start = date(2026, 9, 13)

    print("=" * 110)
    print("TESTE - QUERY 04 - LINE LIST SEMANAL")
    print("=" * 110)
    print(f"WeekStart : {week_start}")
    print()

    rows = execute_query_file(
        BASE_DIR / "sql" / "04_line_list.sql",
        params=(week_start,),
    )

    print(f"Registos devolvidos: {len(rows)}")
    print()

    if not rows:
        print("ATENÇÃO: nenhum registo encontrado.")
        return

    columns = list(rows[0].keys())

    print("=" * 110)
    print("COLUNAS DEVOLVIDAS")
    print("=" * 110)
    for index, column in enumerate(columns, start=1):
        print(f"{index:>2}. {column}")

    print()
    print("=" * 110)
    print("AMOSTRA DOS PRIMEIROS 5 REGISTOS")
    print("=" * 110)

    for index, row in enumerate(rows[:5], start=1):
        print()
        print("-" * 110)
        print(f"REGISTO {index}")
        print("-" * 110)

        for column in columns:
            print(f"{column:<45}: {row.get(column)}")

    print()
    print("=" * 110)
    print("VALIDAÇÕES")
    print("=" * 110)

    disa_column = next(
        (
            name
            for name in ("CodigoDISA", "Código do DISA", "Codigo do DISA")
            if name in columns
        ),
        None,
    )

    if disa_column:
        missing_disa = sum(1 for row in rows if not row.get(disa_column))
        print(f"Código DISA em falta : {missing_disa}/{len(rows)}")
    else:
        raise AssertionError("Coluna Código do DISA não encontrada.")

    allowed_culture_values = {
        None,
        "CMTB",
        "MNT",
        "Negativo",
        "Contaminado",
        "Não realizado",
        "Contaminated - Possible NTM",
    }

    for field in (
        "Resultado Final Cultura Líquida (MGIT)",
        "Resultado Final Cultura Sólida (LJ)",
    ):
        unexpected = {
            row.get(field)
            for row in rows
            if row.get(field) not in allowed_culture_values
        }
        if unexpected:
            print(
                f"ATENÇÃO: valores não normalizados em {field}: "
                f"{sorted(str(v) for v in unexpected)}"
            )

    delamanida_values = {
        row.get("TSA2 - Delamanida")
        for row in rows
    }
    if delamanida_values != {None}:
        raise AssertionError(
            "TSA2 - Delamanida deve estar vazio em todos os registos."
        )

    print(f"Total de registos    : {len(rows)}")
    print()
    print("=" * 110)
    print("PASSOU: Query 04 executada e validada pelo Python.")
    print("=" * 110)


if __name__ == "__main__":
    main()
