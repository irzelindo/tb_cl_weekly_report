"""Verifica contabilização das rejeições semanais sem dados nominais."""
import argparse
from datetime import datetime
from pathlib import Path
from collections import defaultdict

from database import execute_query_file

BASE_DIR = Path(__file__).resolve().parent.parent
LABS = ("Carmelo", "Nampula", "Maputo", "Beira")


def main():
    parser = argparse.ArgumentParser(description="Auditar motivos de rejeição")
    parser.add_argument("--week-start", default="2026-09-13")
    args = parser.parse_args()
    week_start = datetime.strptime(args.week_start, "%Y-%m-%d").date()
    national = defaultdict(int)
    national_total = 0

    for lab in LABS:
        fluxo = execute_query_file(
            BASE_DIR / "sql" / "05_fluxo_pendencias.sql",
            params=(week_start, lab),
        )
        motives = execute_query_file(
            BASE_DIR / "sql" / "06_motivos_rejeicao.sql",
            params=(week_start, lab),
        )
        expected = next((r["Total"] for r in fluxo if r["Codigo"] == "Rejeitadas"), 0)
        actual = sum(int(r["Quantidade"]) for r in motives)
        print(f"{lab:<12} | Fluxo={expected:>6} | Motivos={actual:>6}")
        if expected != actual:
            raise AssertionError(
                f"Rejeições divergentes em {lab}: fluxo={expected}, motivos={actual}"
            )
        for r in motives:
            national[str(r["MotivoAgrupado"])] += int(r["Quantidade"])
        national_total += actual

    print(f"\nTotal de rejeições registadas na semana: {national_total}")
    print("Este número NÃO equivale aos 925 registos históricos sem filtro de datas.\n")
    for reason, n in sorted(national.items(), key=lambda r: (-r[1], r[0])):
        print(f"{reason:<48} {n:>7}")
    print("\nATENÇÃO: COM/Comentário requer validação antes de ser classificado")
    print("como rejeição clínica efectiva. Não divulgar totais sem verificar.")


if __name__ == "__main__":
    main()
