"""Preenchimento do template multilaboratório (7 folhas).

SQL devolve uma população por laboratório. Não há AutoFit e não são
calculados totais de pacientes pela soma de laboratórios.
"""
from __future__ import annotations

import io
import os
import tempfile
from copy import copy
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.cell.cell import MergedCell

from excel_writer import (
    encrypt_workbook, validate_saved_workbook, validate_encrypted_workbook
)

LABS = ("Carmelo", "Nampula", "Maputo", "Beira")
METRIC_CODES = (
    "Recebidas", "Testadas", "Validadas", "Rejeitadas",
    "Pendentes", "Contaminadas", "RR/MR", "Pré-XR", "XR"
)
CULTURE_CODES = (
    "MGIT_Registadas", "MGIT_Testadas", "MGIT_Validadas",
    "MGIT_Rejeitadas", "MGIT_Contaminadas", "MGIT_Pendentes",
    "MGIT_CMTB", "MGIT_MNT", "MGIT_Negativo",
    "LJ_Registadas", "LJ_Testadas", "LJ_Validadas",
    "LJ_Rejeitadas", "LJ_Contaminadas", "LJ_Pendentes",
    "LJ_CMTB", "LJ_MNT", "LJ_Negativo"
)
PRIORITY_ORDER = {"RR/MR": 3, "Pré-XR": 4, "XR": 5}
FIELD_TO_COL = {
    "Código do DISA": 2,
    "Nome do Paciente": 3,
    "Idade": 4,
    "Sexo": 5,
    "US de Colheita": 6,
    "Distrito / Província": 7,
    "Tipo de Amostra": 8,
    "Data de Colheita": 9,
    "Data de Registo no Lab": 10,
    "Data de Processamento": 11,
    "Data de Validação": 12,
    "TRL Registo -> Validação (dias)": 13,
    "Resultado Final Cultura Líquida (MGIT)": 14,
    "Resultado Final Cultura Sólida (LJ)": 15,
    "LPA1 - Isoniazida": 16,
    "LPA1 - Rifampicina": 17,
    "LPA2 - Drogas": 18,
    "LPA CM - Identificação": 19,
    "TSA1 - Streptomicina": 20,
    "TSA1 - Isoniazida": 21,
    "TSA1 - Rifampicina": 22,
    "TSA1 - Etambutol": 23,
    "TSA2 - Bedaquilina": 24,
    "TSA2 - Linezolida": 25,
    "TSA2 - Moxifloxacina": 26,
    "TSA2 - Delamanida": 27,
    "TSA2 - Clofazimina": 28,
    "Classificação da Resistência": 29,
    "Link do Laudo": 30,
}
DATE_FIELDS = {
    "Data de Colheita", "Data de Registo no Lab",
    "Data de Processamento", "Data de Validação"
}
THIN = Side(style="thin", color="BFBFBF")


def safe_write(ws, row, col, value, fmt=None):
    cell = ws.cell(row, col)
    if isinstance(cell, MergedCell):
        raise ValueError("Tentativa de escrita numa célula mesclada secundária")
    cell.value = value
    if fmt is not None:
        cell.number_format = fmt


def integer_day(value):
    """Arredondamento half-up para dias não negativos, sem casas decimais."""
    if value is None:
        return None
    from decimal import Decimal, ROUND_HALF_UP
    return int(Decimal(str(value)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def as_fraction(value):
    return float(value) / 100.0 if value is not None else None


def records_by_order(rows):
    return {int(r["Ordem"]): r for r in rows}


def rows_by_code(rows):
    return {str(r["Codigo"]): r for r in rows}


def metric_record(data, lab, metric):
    block = data[lab]
    if metric in PRIORITY_ORDER:
        return records_by_order(block["indicadores"]).get(PRIORITY_ORDER[metric])
    return rows_by_code(block["fluxo"]).get(metric)


def metric_value(data, lab, metric, field="Total"):
    record = metric_record(data, lab, metric)
    return record.get(field) if record is not None else None


def culture_value(data, lab, code):
    if code in CULTURE_CODES and code.endswith(("_CMTB", "_MNT", "_Negativo")):
        record = records_by_order(data[lab]["indicadores"]).get(1)
        if record is None:
            return None
        return record.get(code)
    record = rows_by_code(data[lab]["fluxo"]).get(code)
    return record.get("Total") if record is not None else None


def write_metadata(wb, week_start, week_end):
    ws = wb["Resumo"]
    safe_write(ws, 4, 2, f"{week_start:%d/%m/%Y} a {week_end:%d/%m/%Y}")
    safe_write(ws, 4, 4, week_start.strftime("%d%m%Y"), "@")
    safe_write(ws, 5, 2, datetime.now().strftime("%d/%m/%Y"))


def write_comparisons(wb, data):
    resumo, comp = wb["Resumo"], wb["Comparação"]
    for i, metric in enumerate(METRIC_CODES):
        for j, lab in enumerate(LABS, start=2):
            value = metric_value(data, lab, metric)
            safe_write(resumo, 10 + i, j, value, "0")
            safe_write(comp, 5 + i, j, value, "0")
        vals = [metric_value(data, lab, metric) for lab in LABS]
        total = sum(vals) if all(v is not None for v in vals) else None
        safe_write(resumo, 10 + i, 6, total, "0")
        safe_write(comp, 5 + i, 6, total, "0")

    # Resultado final da cultura e fluxo MGIT/LJ.
    for i, metric in enumerate(CULTURE_CODES):
        for j, lab in enumerate(LABS, start=2):
            safe_write(resumo, 22 + i, j, culture_value(data, lab, metric), "0")
        vals = [culture_value(data, lab, metric) for lab in LABS]
        safe_write(resumo, 22 + i, 6,
                   sum(vals) if all(v is not None for v in vals) else None, "0")

    # Comparação mantém os indicadores laboratoriais anteriormente existentes.
    extra = ("MGIT_CMTB", "MGIT_MNT", "MGIT_Negativo",
             "LJ_CMTB", "LJ_MNT", "LJ_Negativo",
             "LPA1_Total", "LPA2", "LPA_CM", "TSA_Total")
    for i, metric in enumerate(extra, start=14):
        for j, lab in enumerate(LABS, start=2):
            base = records_by_order(data[lab]["indicadores"]).get(1, {})
            value = base.get(metric)
            safe_write(comp, i, j, value, "0")
        values = [records_by_order(data[lab]["indicadores"]).get(1, {}).get(metric)
                  for lab in LABS]
        safe_write(comp, i, 6, sum(values) if all(v is not None for v in values) else None, "0")


def write_sex(wb, data):
    ws = wb["Desagregação Sexo"]
    for lab_index, lab in enumerate(LABS):
        for metric_index, metric in enumerate(METRIC_CODES):
            row = 5 + lab_index * len(METRIC_CODES) + metric_index
            for col, field in ((4, "Homens"), (5, "Mulheres"),
                               (6, "NaoEspecificado"), (7, "Total")):
                safe_write(ws, row, col, metric_value(data, lab, metric, field), "0")


def write_trl(wb, data):
    ws = wb["TRL"]
    # Linhas 5..20: 4 exames por laboratório, conforme template aprovado.
    # TSA1/2 (quinto indicador anterior) é preservado no fim, linhas 21..24.
    for lab_idx, lab in enumerate(LABS):
        by_order = records_by_order(data[lab]["trl"])
        for order in range(1, 6):
            row = 5 + lab_idx * 4 + order - 1 if order <= 4 else 21 + lab_idx
            if order == 5:
                for col in range(1, 10):
                    source = ws.cell(20, col)
                    target = ws.cell(row, col)
                    if source.has_style:
                        target._style = copy(source._style)
                    target.alignment = copy(source.alignment)
                safe_write(ws, row, 1, lab)
                safe_write(ws, row, 2, "TSA 1 e 2")
                safe_write(ws, row, 5, 30, "0")
            record = by_order.get(order)
            if record is None:
                continue
            safe_write(ws, row, 3, record.get("TotalRegistos"), "0")
            safe_write(ws, row, 4, integer_day(record.get("Media do tempo avaliado para o Target")), "0")
            safe_write(ws, row, 6, record.get("Registos fora do prazo"), "0")
            safe_write(ws, row, 7, as_fraction(record.get("Percentagem fora do prazo")), "0.00%")
            safe_write(ws, row, 8, record.get("Cumprimento do Target pela Media"))
            safe_write(ws, row, 9, record.get("Alerta de Atrasos"))


def write_backlog(wb, data):
    ws = wb["Pendências"]
    # O título original "44-60" sobrepõe 8-44. Ajustado 45-60.
    safe_write(ws, 4, 5, "45–60 dias")
    for i, lab in enumerate(LABS, start=5):
        pending = rows_by_code(data[lab]["fluxo"]).get("Pendentes", {})
        total = pending.get("Total")
        aged = pending.get("Mais60")
        for col, value in {
            2: total, 3: pending.get("0_7"), 4: pending.get("8_44"),
            5: pending.get("45_60"), 6: aged,
            7: pending.get("MaisAntigaDias"),
            8: (aged / total) if (aged is not None and total) else None,
        }.items():
            safe_write(ws, i, col, value, "0.00%" if col == 8 else "0")


def write_resistance(wb, data):
    ws = wb["Resistência"]
    for lab_idx, lab in enumerate(LABS):
        by_order = records_by_order(data[lab]["resistencia"])
        for order in range(1, 11):
            row = 5 + lab_idx * 10 + order - 1
            record = by_order.get(order, {})
            for col, field in ((3, "LPA1"), (4, "LPA2"),
                               (5, "TSA1"), (6, "TSA2"),
                               (7, "TotalResistente"), (8, "NumeroPacientes"),
                               (9, "NumeroTestes")):
                safe_write(ws, row, col, record.get(field), "0")
            safe_write(ws, row, 10,
                       as_fraction(record.get("PercentagemResistencia")), "0.00%")


def write_results(wb, data, protection_password):
    ws = wb["Resultados"]
    # Não alterar os freeze panes / sheet views definidos no template.
    row = 3
    for lab in LABS:
        for record in data[lab]["resultados"]:
            safe_write(ws, row, 1, lab)
            for field, col in FIELD_TO_COL.items():
                value = record.get(field)
                if col == 13 and value is None:
                    value = record.get("TAT Registo -> Validação (dias)")
                if col == 13:
                    value = integer_day(value)
                safe_write(ws, row, col, value,
                           "dd/mm/yyyy" if field in DATE_FIELDS else
                           "0" if col == 13 else None)
                if col == 30 and value:
                    cell = ws.cell(row, col)
                    cell.hyperlink = str(value)
                    cell.font = Font(name="Cambria", size=10,
                                     color="0563C1", underline="single")
            for col in range(1, 31):
                cell = ws.cell(row, col)
                if col != 30 or not cell.hyperlink:
                    f = copy(cell.font)
                    f.name, f.sz = "Cambria", 10
                    cell.font = f
                cell.alignment = Alignment(vertical="center",
                                            horizontal="left" if col in (3,6,7,8,18,19,29,30) else "center",
                                            wrap_text=True)
                cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
            row += 1
    last = max(3, row - 1)
    ws.auto_filter.ref = f"A2:AD{last}"
    ws.print_area = f"A1:AD{last}"
    # Protecção contra edição, diferente da encriptação do ficheiro.
    ws.protection.sheet = True
    ws.protection.set_password(protection_password)
    ws.protection.selectLockedCells = False
    ws.protection.selectUnlockedCells = False
    ws.protection.autoFilter = False
    return last


def generate_multilab_excel(template_path, output_path, week_start, week_end,
                            data, report_password, sheet_password):
    if not report_password:
        raise ValueError("REPORT_PASSWORD tem de estar configurada no .env.")
    template_path, output_path = Path(template_path), Path(output_path)
    wb = load_workbook(template_path)
    required = {"Resumo","Comparação","TRL","Pendências",
                "Resistência","Resultados","Desagregação Sexo"}
    if not required.issubset(set(wb.sheetnames)):
        raise ValueError(f"Folhas em falta: {required - set(wb.sheetnames)}")
    if not all(lab in data for lab in LABS):
        raise ValueError("O relatório exige dados para os quatro laboratórios.")
    write_metadata(wb, week_start, week_end)
    write_comparisons(wb, data)
    write_sex(wb, data)
    write_trl(wb, data)
    write_backlog(wb, data)
    write_resistance(wb, data)
    write_results(wb, data, sheet_password or report_password)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp = tempfile.NamedTemporaryFile(prefix=".tb_plain_", suffix=".xlsx",
                                       dir=output_path.parent, delete=False)
    temp_path = Path(temp.name)
    temp.close()
    try:
        wb.save(temp_path)
        wb.close()
        validate_saved_workbook(temp_path)
        encrypt_workbook(temp_path, output_path, report_password)
        validate_encrypted_workbook(output_path, report_password)
    finally:
        temp_path.unlink(missing_ok=True)
    return output_path
