import argparse
import os
from datetime import datetime, timedelta
from pathlib import Path

from dotenv import load_dotenv

from database import execute_query_file
from excel_writer import generate_excel


BASE_DIR = Path(__file__).resolve().parent.parent

load_dotenv(BASE_DIR / ".env")


def main():
    parser = argparse.ArgumentParser(
        description="Gerar relatório semanal de TB"
    )

    parser.add_argument(
        "--week-start",
        required=True,
        help="Data inicial da semana YYYY-MM-DD",
    )

    parser.add_argument(
        "--no-password",
        action="store_true",
        help=(
            "Gera o relatório sem password. "
            "Usar apenas para validação técnica."
        ),
    )

    args = parser.parse_args()

    week_start = datetime.strptime(
        args.week_start,
        "%Y-%m-%d",
    ).date()

    week_end_exclusive = (
        week_start
        + timedelta(days=7)
    )

    week_end_display = (
        week_end_exclusive
        - timedelta(days=1)
    )

    print("=" * 90)
    print("RELATÓRIO SEMANAL DE TB")
    print("=" * 90)

    print(
        f"Período: "
        f"{week_start:%d/%m/%Y} "
        f"a "
        f"{week_end_display:%d/%m/%Y}"
    )

    print()

    # =========================================================
    # QUERY 1 — INDICADORES
    # =========================================================

    print(
        "Executando Query 1 - Indicadores..."
    )

    indicadores = execute_query_file(
        BASE_DIR / "sql" / "01_indicadores.sql",
        params=(week_start,),
    )

    print(
        f"  {len(indicadores)} linhas"
    )

    # =========================================================
    # QUERY 2 — TAT
    # =========================================================

    print(
        "Executando Query 2 - TAT..."
    )

    tat = execute_query_file(
        BASE_DIR / "sql" / "02_tat.sql",
        params=(week_start,),
    )

    print(
        f"  {len(tat)} linhas"
    )

    # =========================================================
    # QUERY 3 — RESISTÊNCIA
    # =========================================================

    print(
        "Executando Query 3 - Resistência..."
    )

    resistencia = execute_query_file(
        BASE_DIR / "sql" / "03_resistencia.sql",
        params=(week_start,),
    )

    print(
        f"  {len(resistencia)} linhas"
    )

    # =========================================================
    # QUERY 4 — RESULTADOS
    # =========================================================

    print(
        "Executando Query 4 - Resultados..."
    )

    resultados = execute_query_file(
        BASE_DIR / "sql" / "04_line_list.sql",
        params=(week_start,),
    )

    print(
        f"  {len(resultados)} linhas"
    )

    # =========================================================
    # TEMPLATE
    # =========================================================

    template_env = os.getenv(
        "REPORT_TEMPLATE",
        "templates/TB_Weekly_Report_Template.xlsx",
    )

    template_path = Path(
        template_env
    )

    if not template_path.is_absolute():
        template_path = (
            BASE_DIR
            / template_path
        )

    if not template_path.exists():
        raise FileNotFoundError(
            f"Template não encontrado: {template_path}"
        )

    # =========================================================
    # OUTPUT
    # =========================================================

    output_dir_env = os.getenv(
        "REPORT_OUTPUT_DIR",
        "reports",
    )

    output_dir = Path(
        output_dir_env
    )

    if not output_dir.is_absolute():
        output_dir = (
            BASE_DIR
            / output_dir
        )

    output_dir = (
        output_dir
        / str(week_start.year)
    )

    filename = (
        "Relatorio_TB_"
        f"{week_start:%Y-%m-%d}"
        "_a_"
        f"{week_end_display:%Y-%m-%d}"
        ".xlsx"
    )

    output_path = (
        output_dir
        / filename
    )

    # =========================================================
    # GERAR EXCEL
    # =========================================================

    print()
    print("Gerando Excel...")

    generate_excel(
        template_path=template_path,
        output_path=output_path,
        week_start=week_start,
        week_end=week_end_display,
        indicadores=indicadores,
        tat=tat,
        resistencia=resistencia,
        resultados=resultados,
        template_password=os.getenv(
            "REPORT_TEMPLATE_PASSWORD"
        ),
        protect_output=not args.no_password,
    )

    print()
    print("=" * 90)
    print("RELATÓRIO GERADO COM SUCESSO")
    print("=" * 90)
    print()
    print(output_path)

    if args.no_password:
        print(
            "Modo de validação: relatório gerado SEM password."
        )
    else:
        print(
            "O relatório final foi protegido com a password rotativa."
        )


if __name__ == "__main__":
    main()
