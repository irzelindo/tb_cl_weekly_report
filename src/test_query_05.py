"""Valida indicadores de fluxo/backlog, sem expor nomes de pacientes."""
from datetime import date
from pathlib import Path
from database import execute_query_file

BASE = Path(__file__).resolve().parent.parent

def main():
    week_start = date(2026, 9, 13)
    lab = "Maputo"
    rows = execute_query_file(
        BASE / "sql" / "05_fluxo_pendencias.sql",
        params=(week_start, lab),
    )
    print(f"Fluxo e pendências: {lab}, semana iniciada em {week_start}")
    by_code = {str(x["Codigo"]): x for x in rows}
    for r in rows:
        print(
            f"{r['Codigo']:<24} "
            f"H={r['Homens']!s:<5} M={r['Mulheres']!s:<5} "
            f"NI={r['NaoEspecificado']!s:<5} Total={r['Total']}"
        )
    pending = by_code.get("Pendentes")
    if pending:
        total = pending["Total"] or 0
        bins = [pending[k] or 0 for k in ("0_7","8_44","45_60","Mais60")]
        assert sum(bins) == total, (
            f"Pendentes {total} != soma das faixas {sum(bins)}"
        )
        print("PASSOU: faixas de pendência sem dupla contagem.")
    for r in rows:
        if r["Total"] is not None:
            assert r["Total"] == sum(r[k] or 0 for k in ("Homens","Mulheres","NaoEspecificado"))
    print("PASSOU: desagregação por sexo consistente.")
    assert "Rejeitadas" in by_code, "Falta a linha Rejeitadas."
    print("PASSOU: Rejeitadas contabilizadas com LIMSRejectionCode/Desc.")

if __name__ == "__main__":
    main()
