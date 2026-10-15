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



def write_rejections(wb, data, week_start, week_end):
    """Folha compacta de motivos: ranking PT/EN e detalhe por código/descrição.

    O total de cada laboratório tem de conciliar com a Query 05.
    A categoria COM/Comentário continua visível, marcada para revisão.
    """
    from collections import defaultdict
    from openpyxl.styles import PatternFill
    from openpyxl.formatting.rule import DataBarRule
    from openpyxl.utils import get_column_letter

    name = "Rejeições"
    ws = wb[name] if name in wb.sheetnames else wb.create_sheet(name, 4)
    ws.sheet_view.showGridLines = False

    # Conciliar com a população (mesmo período e laboratório) da Query 05.
    groups = defaultdict(lambda: defaultdict(int))
    detailed = []
    for lab in LABS:
        rows = data[lab].get("rejeicoes", [])
        total_reasons = 0
        for rec in rows:
            count = int(rec.get("Quantidade") or 0)
            if count < 0:
                raise ValueError("Quantidade negativa de motivos de rejeição.")
            category = rec.get("MotivoAgrupado") or "Outros / não classificados"
            groups[category][lab] += count
            total_reasons += count
            detailed.append((
                lab, category,
                rec.get("CodigoOriginal") or "(sem código)",
                rec.get("DescricaoOriginal") or "(sem descrição)",
                count,
            ))
        count_main = metric_value(data, lab, "Rejeitadas")
        if count_main is None and not data[lab].get("fluxo"):
            # Não houve registos históricos deste laboratório antes do corte.
            count_main = 0
        if count_main is None or int(count_main) != total_reasons:
            raise ValueError(
                f"Rejeições não conciliadas em {lab}: "
                f"fluxo={count_main}, motivos={total_reasons}. "
                "Validar os filtros antes de divulgar o relatório."
            )

    for row in ws:
        for cell in row:
            if not isinstance(cell, MergedCell):
                cell.value = None

    ws.merge_cells("A1:H2")
    ws["A1"] = "REJEIÇÕES — MOTIVOS EQUIVALENTES POR LABORATÓRIO"
    ws["A1"].fill = PatternFill("solid", fgColor="F4B183")
    ws["A1"].font = Font(name="Cambria", bold=True, size=14, color="242424")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A3:H3")
    ws["A3"] = (
        f"Registos de {week_start:%d/%m/%Y} a {week_end:%d/%m/%Y} | "
        "uma rejeição por amostra; COM/Comentário requer validação."
    )
    ws["A3"].font = Font(name="Cambria", size=10, italic=True)
    ws["A3"].alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[3].height = 28

    summary_headers = [
        "Motivo agrupado", *LABS, "Total", "% do total", "Observação"
    ]
    for col, label in enumerate(summary_headers, 1):
        safe_write(ws, 5, col, label)

    totals = {}
    for group, per_lab in groups.items():
        totals[group] = sum(per_lab.values())
    grand_total = sum(totals.values())
    ordered = sorted(groups, key=lambda g: (-totals[g], g.casefold()))

    for idx, group in enumerate(ordered, start=6):
        safe_write(ws, idx, 1, group)
        for col, lab in enumerate(LABS, start=2):
            safe_write(ws, idx, col, groups[group][lab], "0")
        safe_write(ws, idx, 6, totals[group], "0")
        safe_write(ws, idx, 7, totals[group] / grand_total if grand_total else None, "0.0%")
        safe_write(
            ws, idx, 8,
            "Conferir significado do código no LIMS"
            if ("Comentário" in group or "por validar" in group)
            else ""
        )
        ws.row_dimensions[idx].height = 27

    summary_end = 5 + len(ordered)
    total_row = summary_end + 1
    safe_write(ws, total_row, 1, "TOTAL SEMANAL")
    for col, lab in enumerate(LABS, start=2):
        safe_write(ws, total_row, col, sum(groups[g][lab] for g in ordered), "0")
    safe_write(ws, total_row, 6, grand_total, "0")
    safe_write(ws, total_row, 7, 1 if grand_total else None, "0.0%")

    # Motivos originais para investigação de equivalências entre idiomas.
    details_title = total_row + 3
    ws.merge_cells(
        start_row=details_title, start_column=1,
        end_row=details_title, end_column=8
    )
    ws.cell(details_title, 1, "DETALHE AUDITÁVEL — CÓDIGO E DESCRIÇÃO ORIGINAIS")
    detail_header_row = details_title + 1
    columns = [
        "Laboratório", "Motivo agrupado", "Código original",
        "Descrição original", "Quantidade", "% dentro do laboratório"
    ]
    for col, label in enumerate(columns, start=1):
        safe_write(ws, detail_header_row, col, label)
    detailed.sort(key=lambda r: (LABS.index(r[0]), -r[4], r[1], r[2]))
    lab_totals = {lab: sum(groups[g][lab] for g in ordered) for lab in LABS}
    for row_no, (lab, group, code, desc, count) in enumerate(
        detailed, start=detail_header_row + 1
    ):
        values = (
            lab, group, code, desc, count,
            count / lab_totals[lab] if lab_totals[lab] else None
        )
        for col, value in enumerate(values, 1):
            safe_write(ws, row_no, col, value, "0.0%" if col == 6 else
                       "0" if col == 5 else None)

    light_border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    header_fill = PatternFill("solid", fgColor="F4B183")
    total_fill = PatternFill("solid", fgColor="D9E2F3")
    warning_fill = PatternFill("solid", fgColor="FFF2CC")
    for hdr in (5, detail_header_row):
        for cell in ws[hdr][:8 if hdr == 5 else 6]:
            cell.font = Font(name="Cambria", size=10, bold=True)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = light_border
        ws.row_dimensions[hdr].height = 34

    for row_num in range(6, total_row + 1):
        for col in range(1, 9):
            cell = ws.cell(row_num, col)
            cell.font = Font(name="Cambria", size=10, bold=(row_num == total_row))
            cell.border = light_border
            cell.alignment = Alignment(
                horizontal="left" if col in (1, 8) else "center",
                vertical="center", wrap_text=True
            )
            if row_num == total_row:
                cell.fill = total_fill
            elif "Comentário" in str(ws.cell(row_num, 1).value):
                cell.fill = warning_fill

    for row_num in range(detail_header_row + 1, detail_header_row + 1 + len(detailed)):
        for col in range(1, 7):
            cell = ws.cell(row_num, col)
            cell.font = Font(name="Cambria", size=10)
            cell.border = light_border
            cell.alignment = Alignment(
                horizontal="left" if col in (1,2,3,4) else "center",
                vertical="center", wrap_text=True
            )
    for col, width in dict(A=38, B=15, C=15, D=39, E=14, F=15, G=14, H=39).items():
        ws.column_dimensions[col].width = width

    if ordered:
        ws.conditional_formatting.add(
            f"F6:F{summary_end}",
            DataBarRule(start_type="num", start_value=0,
                        end_type="max", color="F4B183", showValue=True)
        )
    if detailed:
        ws.auto_filter.ref = (
            f"A{detail_header_row}:F{detail_header_row + len(detailed)}"
        )
    ws.freeze_panes = "B6"
    ws.print_options.horizontalCentered = True
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    return grand_total

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
    write_rejections(wb, data, week_start, week_end)
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
