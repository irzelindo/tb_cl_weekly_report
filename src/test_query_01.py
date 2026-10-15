from datetime import date
from pathlib import Path

from database import execute_query_file

BASE_DIR = Path(__file__).resolve().parent.parent


def main():
    week_start = date(2026, 9, 13)

    print("=" * 100)
    print("TESTE - QUERY 01 - INDICADORES SEMANAIS")
    print("=" * 100)
    print()
    print(f"WeekStart : {week_start}")
    print()

    rows = execute_query_file(
        BASE_DIR / "sql" / "01_indicadores.sql",
        params=(week_start,),
    )

    if not rows:
        raise RuntimeError("A query não devolveu resultados.")

    print(f"Linhas devolvidas: {len(rows)}")
    print()

    columns = [
        "Indicador",
        "Homens",
        "Mulheres",
        "NaoEspecificado",
        "Liquido_MGIT",
        "Solido_LJ",
        "MGIT_CMTB",
        "MGIT_MNT",
        "MGIT_Negativo",
        "LJ_CMTB",
        "LJ_MNT",
        "LJ_Negativo",
        "LPA1_CMTB",
        "LPA1_MNT",
        "LPA1_Total",
        "LPA2",
        "LPA_CM",
        "TSA1",
        "TSA2",
        "TSA_Total",
        "Total",
    ]

    for index, row in enumerate(rows, start=1):
        print("-" * 100)
        print(f"Linha {index}")
        print("-" * 100)

        for column in columns:
            print(f"{column:<30}: {row.get(column)}")

        print()

    print("=" * 100)
    print("VALIDAÇÕES DE CONSISTÊNCIA")
    print("=" * 100)

    errors = []

    if len(rows) != 5:
        errors.append(
            f"Esperadas 5 linhas, recebidas {len(rows)}."
        )

    base = next(
        (row for row in rows if row.get("Ordem") == 1),
        None,
    )

    if base is None:
        errors.append("Não foi encontrada a linha Ordem=1.")
    else:
        lpa1_total = base.get("LPA1_Total") or 0
        lpa1_cmtb = base.get("LPA1_CMTB") or 0
        lpa1_mnt = base.get("LPA1_MNT") or 0

        if lpa1_cmtb > lpa1_total:
            errors.append(
                f"LPA1_CMTB ({lpa1_cmtb}) > LPA1_Total ({lpa1_total})."
            )

        if lpa1_mnt > lpa1_total:
            errors.append(
                f"LPA1_MNT ({lpa1_mnt}) > LPA1_Total ({lpa1_total})."
            )

        mgit_classificados = sum(
            base.get(field) or 0
            for field in (
                "MGIT_CMTB",
                "MGIT_MNT",
                "MGIT_Negativo",
            )
        )

        if mgit_classificados > (base.get("Liquido_MGIT") or 0):
            errors.append(
                "Resultados finais MGIT excedem MGIT processados."
            )

        lj_classificados = sum(
            base.get(field) or 0
            for field in (
                "LJ_CMTB",
                "LJ_MNT",
                "LJ_Negativo",
            )
        )

        if lj_classificados > (base.get("Solido_LJ") or 0):
            errors.append(
                "Resultados finais LJ excedem LJ processados."
            )

        tsa_total = base.get("TSA_Total") or 0
        if tsa_total < max(
            base.get("TSA1") or 0,
            base.get("TSA2") or 0,
        ):
            errors.append(
                "TSA_Total é inferior a TSA1 ou TSA2."
            )

    if errors:
        print("FALHOU:")
        for error in errors:
            print(f" - {error}")
        raise SystemExit(1)

    print("PASSOU: estrutura e invariantes principais estão consistentes.")
    print("=" * 100)


if __name__ == "__main__":
    main()
