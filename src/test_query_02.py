from datetime import date
from pathlib import Path

from database import execute_query_file

BASE_DIR = Path(__file__).resolve().parent.parent


def main():
    week_start = date(2026, 9, 13)

    print("=" * 110)
    print("TESTE - QUERY 02 - TAT SEMANAL")
    print("=" * 110)
    print(f"WeekStart : {week_start}")
    print()

    rows = execute_query_file(
        BASE_DIR / "sql" / "02_tat.sql",
        params=(week_start,),
    )

    if not rows:
        print("A query não devolveu resultados.")
        return

    print(f"Linhas devolvidas: {len(rows)}")
    print()

    columns = [
        "DataInicio",
        "DataFim",
        "Ordem",
        "Exame",
        "TotalRegistos",
        "RegistosComTempoValido",
        "Registos sem tempo valido",
        "Media Colheita ao Registo",
        "Media Registo ao Processamento",
        "Media Processamento a Validacao",
        "Media Total Colheita a Validacao",
        "Target",
        "Media do tempo avaliado para o Target",
        "Maior tempo para o Target",
        "Registos fora do prazo",
        "Percentagem fora do prazo",
        "Cumprimento do Target pela Media",
        "Alerta de Atrasos",
    ]

    for index, row in enumerate(rows, start=1):
        print("-" * 110)
        print(f"Linha {index}")
        print("-" * 110)

        for column in columns:
            print(f"{column:<42}: {row.get(column)}")
        print()

    expected = {
        1: "Cultura Líquida (MGIT)",
        2: "Cultura Sólida (LJ)",
        3: "LPA - Amostras de Expectoração",
        4: "LPA - Cultura Positiva",
        5: "TSA 1 e 2",
    }

    by_order = {row["Ordem"]: row for row in rows}

    # É aceitável uma categoria não aparecer se não houver exames nessa semana,
    # mas as linhas de cultura devem existir se houve resultados MGIT/LJ.
    for order, name in expected.items():
        if order in by_order and by_order[order]["Exame"] != name:
            raise AssertionError(
                f"Ordem {order}: esperado '{name}', recebido '{by_order[order]['Exame']}'."
            )

    if 1 in by_order and "44 dias" not in str(by_order[1].get("Target")):
        raise AssertionError("Target MGIT deve ser 44 dias.")

    if 2 in by_order and not any(
        text in str(by_order[2].get("Target"))
        for text in ("60 dias", "62 dias")
    ):
        raise AssertionError("Target LJ deve ser 60 ou 62 dias.")

    print("=" * 110)
    print(f"PASSOU: Query 02 devolveu {len(rows)} categorias de TAT.")
    print("=" * 110)


if __name__ == "__main__":
    main()
