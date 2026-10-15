"""Gerador semanal TB — 4 laboratórios, TRL, backlog histórico e Excel encriptado."""
import argparse
import os
from datetime import datetime, timedelta
from pathlib import Path

from dotenv import load_dotenv
from database import execute_query_file
from multilab_writer import LABS, generate_multilab_excel

BASE_DIR = Path(__file__).resolve().parent.parent
load_dotenv(BASE_DIR / ".env")

QUERIES = {
    "indicadores": "01_indicadores.sql",
    "trl": "02_tat.sql",  # nome físico mantido por compatibilidade
    "resistencia": "03_resistencia.sql",
    "resultados": "04_line_list.sql",
    "fluxo": "05_fluxo_pendencias.sql",
}


def main():
    parser = argparse.ArgumentParser(
        description="Relatório semanal TB: Carmelo, Nampula, Maputo e Beira"
    )
    parser.add_argument("--week-start", required=True,
                        help="Início da semana no formato YYYY-MM-DD")
    args = parser.parse_args()
    week_start = datetime.strptime(args.week_start, "%Y-%m-%d").date()
    week_end = week_start + timedelta(days=6)

    report_password = os.getenv("REPORT_PASSWORD")
    if not report_password:
        raise RuntimeError(
            "Defina REPORT_PASSWORD no .env (senha estática de abertura do Excel)."
        )
    sheet_password = os.getenv("REPORT_SHEET_PASSWORD") or report_password

    template_setting = os.getenv(
        "REPORT_TEMPLATE", "templates/TB_Weekly_Report_Template.xlsx"
    )
    template = Path(template_setting)
    if not template.is_absolute():
        template = BASE_DIR / template
    if not template.is_file():
        raise FileNotFoundError(f"Template não encontrado: {template}")

    output_setting = os.getenv("REPORT_OUTPUT_DIR", "reports")
    output_dir = Path(output_setting)
    if not output_dir.is_absolute():
        output_dir = BASE_DIR / output_dir
    output = (
        output_dir / str(week_start.year)
        / f"Relatorio_TB_{week_start:%Y-%m-%d}_a_{week_end:%Y-%m-%d}.xlsx"
    )

    dados = {}
    print(f"RELATÓRIO TB — {week_start:%d/%m/%Y} a {week_end:%d/%m/%Y}")
    for lab in LABS:
        print(f"Laboratório: {lab}")
        block = {}
        for key, filename in QUERIES.items():
            rows = execute_query_file(
                BASE_DIR / "sql" / filename, params=(week_start, lab)
            )
            block[key] = rows
            print(f"  {key}: {len(rows)} registo(s)")
        dados[lab] = block

    generate_multilab_excel(
        template_path=template,
        output_path=output,
        week_start=week_start,
        week_end=week_end,
        data=dados,
        report_password=report_password,
        sheet_password=sheet_password,
    )
    print(f"Relatório gerado e encriptado: {output}")
    print("A senha foi lida do .env e não será apresentada nos logs.")


if __name__ == "__main__":
    main()
