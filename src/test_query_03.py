from datetime import date
from pathlib import Path

from database import execute_query_file

BASE_DIR = Path(__file__).resolve().parent.parent


def main():
    week_start = date(2026, 9, 13)

    print("=" * 110)
    print("TESTE - QUERY 03 - RESISTÊNCIA AOS MEDICAMENTOS")
    print("=" * 110)
    print(f"WeekStart : {week_start}")
    print()

    rows = execute_query_file(
        BASE_DIR / "sql" / "03_resistencia.sql",
        params=(week_start,),
    )

    if not rows:
        print("A query não devolveu resultados.")
        return

    print(f"Linhas devolvidas: {len(rows)}")
    print()

    columns = [
        "Medicamento",
        "LPA1",
        "LPA2",
        "TSA1",
        "TSA2",
        "TotalResistente",
        "NumeroPacientes",
        "NumeroTestes",
        "PercentagemResistencia",
    ]

    for index, row in enumerate(rows, start=1):
        print("-" * 110)
        print(f"Linha {index}")
        print("-" * 110)

        for column in columns:
            print(f"{column:<30}: {row.get(column)}")
        print()

        pacientes = row.get("NumeroPacientes")
        testes = row.get("NumeroTestes")
        resistentes = row.get("TotalResistente")
        percentagem = row.get("PercentagemResistencia")

        if pacientes is not None and testes is not None and testes < pacientes:
            raise AssertionError(
                f"{row['Medicamento']}: NumeroTestes < NumeroPacientes."
            )

        if pacientes is not None and resistentes is not None and resistentes > pacientes:
            raise AssertionError(
                f"{row['Medicamento']}: TotalResistente > NumeroPacientes."
            )

        if percentagem is not None and not (0 <= float(percentagem) <= 100):
            raise AssertionError(
                f"{row['Medicamento']}: percentagem fora de 0-100."
            )

    if len(rows) != 10:
        raise AssertionError(
            f"Esperadas 10 linhas; recebidas {len(rows)}."
        )

    delamanida = next(row for row in rows if row["Medicamento"] == "Delamanida")
    if any(delamanida.get(field) is not None for field in columns[1:]):
        raise AssertionError(
            "Delamanida deve permanecer vazia porque TBCultura não possui esse campo."
        )

    print("=" * 110)
    print("PASSOU: Query 03 validada, incluindo aplicabilidade dos métodos.")
    print("=" * 110)


if __name__ == "__main__":
    main()
