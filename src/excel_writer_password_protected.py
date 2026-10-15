from copy import copy
from datetime import date, datetime, timedelta
from pathlib import Path
import io
import os
import re
import tempfile
import zipfile

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter


# =============================================================
# CONSTANTES / MAPEAMENTOS — FOLHA RESULTADOS
# =============================================================

RESULTADOS_MAPPING = {
    "Código do DISA": 1,                           # A
    "Nome do Paciente": 2,                         # B
    "Idade": 3,                                    # C
    "Sexo": 4,                                     # D
    "US de Colheita": 5,                           # E
    "Distrito / Província": 6,                     # F
    "Tipo de Amostra": 7,                          # G
    "Data de Colheita": 8,                         # H
    "Data de Registo no Lab": 9,                   # I
    "Data de Processamento": 10,                   # J
    "Data de Validação": 11,                       # K
    "TAT Registo -> Validação (dias)": 12,         # L
    "Resultado Final Cultura Líquida (MGIT)": 13,  # M
    "Resultado Final Cultura Sólida (LJ)": 14,     # N
    "LPA1 - Isoniazida": 15,                       # O
    "LPA1 - Rifampicina": 16,                      # P
    "LPA2 - Drogas": 17,                           # Q (Q:R mesclado)
    "LPA CM - Identificação": 19,                  # S
    "TSA1 - Streptomicina": 20,                    # T
    "TSA1 - Isoniazida": 21,                       # U
    "TSA1 - Rifampicina": 22,                      # V
    "TSA1 - Etambutol": 23,                        # W
    "TSA2 - Bedaquilina": 24,                      # X
    "TSA2 - Linezolida": 25,                       # Y
    "TSA2 - Moxifloxacina": 26,                    # Z
    "TSA2 - Delamanida": 27,                       # AA
    "TSA2 - Clofazimina": 28,                      # AB
    "Classificação da Resistência": 29,            # AC
    "Link do Laudo": 30,                           # AD
}

RESULTADOS_DATE_FIELDS = {
    "Data de Colheita",
    "Data de Registo no Lab",
    "Data de Processamento",
    "Data de Validação",
}


# =============================================================
# FUNÇÕES AUXILIARES
# =============================================================


def copy_cell_style(source, target):
    """Copia estilo sem tentar escrever em células secundárias de merges."""
    if isinstance(source, MergedCell) or isinstance(target, MergedCell):
        return

    if source.has_style:
        target.font = copy(source.font)
        target.fill = copy(source.fill)
        target.border = copy(source.border)
        target.alignment = copy(source.alignment)
        target.number_format = source.number_format
        target.protection = copy(source.protection)



def copy_row_style(ws, source_row, target_row, start_col=1, end_col=30):
    """Copia a formatação de uma linha-modelo."""
    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height

    for col in range(start_col, end_col + 1):
        source = ws.cell(row=source_row, column=col)
        target = ws.cell(row=target_row, column=col)

        if isinstance(source, MergedCell) or isinstance(target, MergedCell):
            continue

        copy_cell_style(source, target)



def normalize_font_range(ws, cell_range, font_name="Cambria", font_size=10):
    """Uniformiza nome e tamanho da fonte preservando o restante estilo."""
    for row in ws[cell_range]:
        for cell in row:
            if isinstance(cell, MergedCell):
                continue

            font = copy(cell.font)
            font.name = font_name
            font.sz = font_size
            cell.font = font



def enable_wrap_text(ws, cell_range):
    """Activa wrap_text preservando o restante alinhamento."""
    for row in ws[cell_range]:
        for cell in row:
            if isinstance(cell, MergedCell):
                continue

            alignment = copy(cell.alignment)
            alignment.wrap_text = True
            cell.alignment = alignment



def autofit_columns(
    ws,
    start_col,
    end_col,
    start_row,
    end_row,
    min_width=8,
    max_width=30,
    padding=2,
    custom_max_widths=None,
    skip_columns=None,
):
    """AutoFit controlado, com limites para evitar colunas excessivamente largas."""
    custom_max_widths = custom_max_widths or {}
    skip_columns = set(skip_columns or [])

    for col_num in range(start_col, end_col + 1):
        col_letter = get_column_letter(col_num)

        if col_letter in skip_columns:
            continue

        max_length = 0

        for row_num in range(start_row, end_row + 1):
            cell = ws.cell(row=row_num, column=col_num)

            if isinstance(cell, MergedCell) or cell.value is None:
                continue

            lines = str(cell.value).splitlines() or [""]
            max_length = max(max_length, max(len(line) for line in lines))

        width = max(min_width, min(max_length + padding, custom_max_widths.get(col_letter, max_width)))
        ws.column_dimensions[col_letter].width = width



def find_section_row(ws, text, max_row=120, max_col=6):
    """Localiza uma secção pelo texto, sem depender de números fixos de linha."""
    wanted = text.lower()

    for row in range(1, min(ws.max_row, max_row) + 1):
        for col in range(1, min(ws.max_column, max_col) + 1):
            cell = ws.cell(row=row, column=col)

            if isinstance(cell, MergedCell) or cell.value is None:
                continue

            if wanted in str(cell.value).lower():
                return row

    raise ValueError(f"Secção não encontrada na folha '{ws.title}': {text}")



def _rgb_from_color(color):
    if color is None:
        return None

    if getattr(color, "type", None) == "rgb" and getattr(color, "rgb", None):
        value = color.rgb[-6:]
        try:
            return tuple(int(value[i:i+2], 16) for i in (0, 2, 4))
        except ValueError:
            return None

    return None



def is_red_highlight(cell):
    """
    Detecta células de dados marcadas a vermelho no template.

    A regra é usada apenas nas linhas de dados. Assim, cabeçalhos salmão
    não interferem. Uma célula marcada a vermelho é tratada como
    "não aplicável" e permanece vazia.
    """
    if isinstance(cell, MergedCell):
        return False

    fill = cell.fill
    if fill is None or fill.fill_type is None:
        return False

    rgb = _rgb_from_color(fill.fgColor)
    if rgb is None:
        return False

    r, g, b = rgb

    # Vermelho forte ou vermelho-claro usado para marcação manual.
    return (
        (r >= 200 and g <= 150 and b <= 150)
        or (r >= 235 and g <= 215 and b <= 215 and r >= g + 20 and r >= b + 20)
    )



def set_template_value(cell, value, number_format=None):
    """Escreve somente se a célula não estiver marcada a vermelho."""
    if isinstance(cell, MergedCell):
        return

    if is_red_highlight(cell):
        cell.value = None
        return

    cell.value = value

    if number_format is not None and value is not None:
        cell.number_format = number_format



def set_hyperlink(cell, url):
    """Cria hyperlink sem aplicar o NamedStyle Hyperlink, que altera a fonte."""
    if is_red_highlight(cell):
        cell.value = None
        cell.hyperlink = None
        return

    cell.hyperlink = url

    font = copy(cell.font)
    font.name = "Cambria"
    font.sz = 10
    font.color = "0563C1"
    font.underline = "single"
    cell.font = font


# =============================================================
# FOLHA RESULTADOS — QUERY 4
# =============================================================


def write_resultados(ws, rows):
    """Preenche o line-list semanal preservando o merge Q:R."""
    DATA_START_ROW = 3
    TEMPLATE_ROW = 3
    existing_last_row = max(ws.max_row, DATA_START_ROW)

    if rows:
        last_data_row = DATA_START_ROW + len(rows) - 1
    else:
        last_data_row = DATA_START_ROW

    # Criar linhas adicionais, se necessário.
    if last_data_row > existing_last_row:
        existing_merges = {str(rng) for rng in ws.merged_cells.ranges}

        for row_num in range(existing_last_row + 1, last_data_row + 1):
            copy_row_style(ws, TEMPLATE_ROW, row_num, 1, 30)

            merge_range = f"Q{row_num}:R{row_num}"
            if merge_range not in existing_merges:
                ws.merge_cells(merge_range)
                existing_merges.add(merge_range)

    clear_until = max(existing_last_row, last_data_row)

    # Limpar dados anteriores sem apagar estilos/cores do template.
    for row_num in range(DATA_START_ROW, clear_until + 1):
        for col_num in list(range(1, 18)) + list(range(19, 31)):
            cell = ws.cell(row=row_num, column=col_num)

            if isinstance(cell, MergedCell):
                continue

            cell.value = None
            cell.hyperlink = None

    for index, record in enumerate(rows):
        excel_row = DATA_START_ROW + index

        for field, column in RESULTADOS_MAPPING.items():
            cell = ws.cell(row=excel_row, column=column)

            if isinstance(cell, MergedCell) or is_red_highlight(cell):
                continue

            value = record.get(field)
            cell.value = value

            if field in RESULTADOS_DATE_FIELDS and value is not None:
                cell.number_format = "dd/mm/yyyy"

            if field == "Link do Laudo" and value:
                set_hyperlink(cell, value)

    if rows:
        ws.auto_filter.ref = f"A2:AD{last_data_row}"
    else:
        ws.auto_filter.ref = "A2:AD3"

    return last_data_row


# =============================================================
# RESUMO — QUERY 1: INDICADORES
# =============================================================


def _indicadores_mapping(ws, title_row):
    """
    Novo layout esperado:

      A Indicador
      B Homens
      C Mulheres
      D:E Tipo de Cultura
      F:H Resultado Final MGIT  -> CMTB / MNT / Negativo
      I:K Resultado Final LJ    -> CMTB / MNT / Negativo
      L:M LPA1                  -> CMTB / MNT
      N:O LPA2 / LPA CM
      P TSA1
      Q TSA2
      R Total

    Se o template incluir "Não Especificado" na coluna D, o mapeamento
    é deslocado automaticamente uma coluna para a direita.
    """
    subheader_row = title_row + 2
    d_value = ws.cell(row=subheader_row, column=4).value
    has_unknown_sex = d_value is not None and (
        "especific" in str(d_value).lower()
        or "não especific" in str(d_value).lower()
    )

    if has_unknown_sex:
        return {
            "Homens": 2,
            "Mulheres": 3,
            "NaoEspecificado": 4,
            "Liquido_MGIT": 5,
            "Solido_LJ": 6,
            "MGIT_CMTB": 7,
            "MGIT_MNT": 8,
            "MGIT_Negativo": 9,
            "LJ_CMTB": 10,
            "LJ_MNT": 11,
            "LJ_Negativo": 12,
            "LPA1_CMTB": 13,
            "LPA1_MNT": 14,
            "LPA2": 15,
            "LPA_CM": 16,
            "TSA1": 17,
            "TSA2": 18,
            "Total": 19,
        }

    return {
        "Homens": 2,
        "Mulheres": 3,
        "Liquido_MGIT": 4,
        "Solido_LJ": 5,
        "MGIT_CMTB": 6,
        "MGIT_MNT": 7,
        "MGIT_Negativo": 8,
        "LJ_CMTB": 9,
        "LJ_MNT": 10,
        "LJ_Negativo": 11,
        "LPA1_CMTB": 12,
        "LPA1_MNT": 13,
        "LPA2": 14,
        "LPA_CM": 15,
        "TSA1": 16,
        "TSA2": 17,
        "Total": 18,
    }



def write_indicadores(ws, rows):
    title_row = find_section_row(ws, "1. Indicadores")
    data_start = title_row + 3
    mapping = _indicadores_mapping(ws, title_row)
    max_col = max(mapping.values())

    rows_by_order = {row["Ordem"]: row for row in rows}

    # Limpar somente valores das 5 linhas da tabela.
    for r in range(data_start, data_start + 5):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            if not isinstance(cell, MergedCell):
                cell.value = None

    for order in range(1, 6):
        excel_row = data_start + order - 1
        record = rows_by_order.get(order)

        if not record:
            continue

        set_template_value(ws.cell(excel_row, 1), record.get("Indicador"))

        for field, col in mapping.items():
            value = record.get(field)
            cell = ws.cell(row=excel_row, column=col)

            # NULL = não aplicável -> vazio, nunca zero nem travessão.
            set_template_value(cell, value, "0" if value is not None else None)

    return title_row, data_start, max_col


# =============================================================
# RESUMO — QUERY 2: TAT
# =============================================================


def write_tat(ws, rows):
    title_row = find_section_row(ws, "2. Tempo de Resposta")
    data_start = title_row + 2

    rows_by_order = {row["Ordem"]: row for row in rows}

    # Novo TAT tem 5 linhas: MGIT, LJ, LPA directo, LPA cultura positiva, TSA.
    for r in range(data_start, data_start + 5):
        for c in range(1, 12):
            cell = ws.cell(row=r, column=c)
            if not isinstance(cell, MergedCell):
                cell.value = None

    for order in range(1, 6):
        excel_row = data_start + order - 1
        record = rows_by_order.get(order)

        if not record:
            continue

        set_template_value(ws.cell(excel_row, 1), record.get("Exame"))

        numeric_fields = {
            2: "Media Colheita ao Registo",
            3: "Media Registo ao Processamento",
            4: "Media Processamento a Validacao",
            5: "Media Total Colheita a Validacao",
            6: "Media do tempo avaliado para o Target",
        }

        for col, field in numeric_fields.items():
            value = record.get(field)
            set_template_value(ws.cell(excel_row, col), value, "0.00" if value is not None else None)

        set_template_value(ws.cell(excel_row, 7), record.get("Target"))
        set_template_value(ws.cell(excel_row, 8), record.get("Registos fora do prazo"), "0")

        percentage = record.get("Percentagem fora do prazo")
        pct_cell = ws.cell(excel_row, 9)
        if percentage is None:
            set_template_value(pct_cell, None)
        else:
            set_template_value(pct_cell, float(percentage) / 100.0, "0.00%")

        set_template_value(ws.cell(excel_row, 10), record.get("Cumprimento do Target pela Media"))
        set_template_value(ws.cell(excel_row, 11), record.get("Alerta de Atrasos"))

    return title_row, data_start


# =============================================================
# RESUMO — QUERY 3: RESISTÊNCIA
# =============================================================


def write_resistencia(ws, rows):
    title_row = find_section_row(ws, "3. Resistência")
    data_start = title_row + 3

    rows_by_order = {row["Ordem"]: row for row in rows}

    for r in range(data_start, data_start + 10):
        for c in range(1, 10):
            cell = ws.cell(row=r, column=c)
            if not isinstance(cell, MergedCell):
                cell.value = None

    fields = {
        2: "LPA1",
        3: "LPA2",
        4: "TSA1",
        5: "TSA2",
        6: "TotalResistente",
        7: "NumeroPacientes",
        8: "NumeroTestes",
    }

    for order in range(1, 11):
        excel_row = data_start + order - 1
        record = rows_by_order.get(order)

        if not record:
            continue

        set_template_value(ws.cell(excel_row, 1), record.get("Medicamento"))

        for col, field in fields.items():
            value = record.get(field)
            # NULL distingue "não aplicável" de zero resistentes.
            set_template_value(ws.cell(excel_row, col), value, "0" if value is not None else None)

        percentage = record.get("PercentagemResistencia")
        pct_cell = ws.cell(excel_row, 9)

        if percentage is None:
            set_template_value(pct_cell, None)
        else:
            set_template_value(pct_cell, float(percentage) / 100.0, "0.00%")

    return title_row, data_start


# =============================================================
# CABEÇALHO / METADADOS
# =============================================================


def write_report_metadata(ws, week_start, week_end):
    ws["B7"] = f"{week_start:%d/%m/%Y} a {week_end:%d/%m/%Y}"

    report_number = week_start.strftime("%d%m%Y")
    ws["D7"] = report_number
    ws["D7"].number_format = "@"

    ws["B8"] = datetime.now().strftime("%d/%m/%Y")

    indicadores_row = find_section_row(ws, "1. Indicadores")
    tat_row = find_section_row(ws, "2. Tempo de Resposta")

    ws.cell(indicadores_row, 1).value = (
        "1. Indicadores de TB — "
        f"{week_start:%d/%m/%Y} a {week_end:%d/%m/%Y}"
    )
    ws.cell(tat_row, 1).value = "2. Tempo de Resposta (TAT) — média em dias"

    return report_number


# =============================================================
# SENHA ROTATIVA / TEMPLATE ENCRIPTADO
# =============================================================

PORTUGUESE_MONTHS = {
    1: "janeiro",
    2: "fevereiro",
    3: "março",
    4: "abril",
    5: "maio",
    6: "junho",
    7: "julho",
    8: "agosto",
    9: "setembro",
    10: "outubro",
    11: "novembro",
    12: "dezembro",
}



def build_report_password(week_start=None):
    """Senha estática definida no .env, independente da semana."""
    password = os.getenv("REPORT_PASSWORD")
    if not password:
        raise RuntimeError("Defina REPORT_PASSWORD no .env antes de gerar o Excel.")
    return password


def _infer_password_from_filename(template_path):
    match = re.search(r"(\d{4})-(\d{2})-(\d{2})", Path(template_path).name)
    if not match:
        return None

    try:
        start = date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
    except ValueError:
        return None

    return build_report_password(start)



def load_template_workbook(template_path, week_start, template_password=None):
    """
    Abre template simples ou encriptado.

    Para automação contínua recomenda-se manter o TEMPLATE sem password e
    encriptar apenas os relatórios finais. Contudo, esta função também aceita
    o template protegido anexado nesta revisão.
    """
    template_path = Path(template_path)

    if zipfile.is_zipfile(template_path):
        return load_workbook(template_path)

    try:
        import msoffcrypto
    except ImportError as exc:
        raise RuntimeError(
            "O template parece estar encriptado. Instale 'msoffcrypto-tool' "
            "ou forneça uma cópia não protegida do template."
        ) from exc

    candidates = []

    for candidate in (
        template_password,
        os.getenv("REPORT_TEMPLATE_PASSWORD"),
        _infer_password_from_filename(template_path),
        build_report_password(week_start),
    ):
        if candidate and candidate not in candidates:
            candidates.append(candidate)

    last_error = None

    for password in candidates:
        try:
            decrypted = io.BytesIO()
            with template_path.open("rb") as source:
                office_file = msoffcrypto.OfficeFile(source)
                office_file.load_key(password=password, verify_password=True)
                office_file.decrypt(decrypted, verify_integrity=True)

            decrypted.seek(0)
            return load_workbook(decrypted)
        except Exception as exc:  # tenta próximo candidato
            last_error = exc

    raise RuntimeError(
        "Não foi possível abrir o template encriptado. Defina "
        "REPORT_TEMPLATE_PASSWORD ou use um template não protegido."
    ) from last_error



def encrypt_workbook(plain_path, encrypted_path, password):
    try:
        from msoffcrypto.format.ooxml import OOXMLFile
    except ImportError as exc:
        raise RuntimeError(
            "A protecção por senha requer 'msoffcrypto-tool'. "
            "Instale com: pip install msoffcrypto-tool"
        ) from exc

    with Path(plain_path).open("rb") as source:
        office_file = OOXMLFile(source)
        with Path(encrypted_path).open("wb") as destination:
            office_file.encrypt(password, destination)



def validate_saved_workbook(path):
    path = Path(path)

    with zipfile.ZipFile(path, "r") as archive:
        bad_file = archive.testzip()
        if bad_file is not None:
            raise ValueError(f"XLSX corrompido: {bad_file}")

    validation_wb = load_workbook(path, read_only=False, data_only=False)
    validation_wb.close()



def validate_encrypted_workbook(path, password):
    try:
        import msoffcrypto
    except ImportError as exc:
        raise RuntimeError("Validação requer 'msoffcrypto-tool'.") from exc

    decrypted = io.BytesIO()

    with Path(path).open("rb") as source:
        office_file = msoffcrypto.OfficeFile(source)

        if not office_file.is_encrypted():
            raise ValueError("O relatório final não está encriptado.")

        office_file.load_key(password=password, verify_password=True)
        office_file.decrypt(decrypted, verify_integrity=True)

    decrypted.seek(0)
    if not zipfile.is_zipfile(decrypted):
        raise ValueError("Conteúdo desencriptado não é um XLSX válido.")


# =============================================================
# GERAÇÃO DO RELATÓRIO
# =============================================================


def generate_excel(
    template_path,
    output_path,
    week_start,
    week_end,
    indicadores,
    tat,
    resistencia,
    resultados,
    template_password=None,
    protect_output=True,
):
    template_path = Path(template_path)
    output_path = Path(output_path)

    if not template_path.exists():
        raise FileNotFoundError(f"Template não encontrado: {template_path}")

    wb = load_template_workbook(
        template_path=template_path,
        week_start=week_start,
        template_password=template_password,
    )

    required_sheets = {"Resumo", "Resultados"}
    missing = required_sheets - set(wb.sheetnames)
    if missing:
        raise KeyError("Folhas em falta no template: " + ", ".join(sorted(missing)))

    ws_resumo = wb["Resumo"]
    ws_resultados = wb["Resultados"]

    write_report_metadata(ws_resumo, week_start, week_end)
    ind_title, ind_start, ind_max_col = write_indicadores(ws_resumo, indicadores)
    tat_title, tat_start = write_tat(ws_resumo, tat)
    res_title, res_start = write_resistencia(ws_resumo, resistencia)
    last_result_row = write_resultados(ws_resultados, resultados)

    # =========================================================
    # FORMATAÇÃO — RESUMO
    # =========================================================
    summary_last_col = max(ind_max_col, 18)
    summary_last_letter = get_column_letter(summary_last_col)
    summary_last_row = res_start + 9

    normalize_font_range(
        ws_resumo,
        f"A{ind_title + 1}:{summary_last_letter}{summary_last_row}",
        "Cambria",
        10,
    )
    enable_wrap_text(
        ws_resumo,
        f"A{ind_title}:{summary_last_letter}{summary_last_row}",
    )

    autofit_columns(
        ws_resumo,
        start_col=1,
        end_col=summary_last_col,
        start_row=ind_title,
        end_row=summary_last_row,
        min_width=8,
        max_width=22,
        custom_max_widths={
            "A": 48,
            "G": 28,
            "J": 20,
            "K": 42,
            "R": 14,
            "S": 14,
        },
    )

    # =========================================================
    # FORMATAÇÃO — RESULTADOS
    # =========================================================
    normalize_font_range(ws_resultados, f"A1:AD{last_result_row}", "Cambria", 10)
    enable_wrap_text(ws_resultados, f"A1:AD{last_result_row}")

    autofit_columns(
        ws_resultados,
        start_col=1,
        end_col=30,
        start_row=1,
        end_row=last_result_row,
        min_width=8,
        max_width=24,
        custom_max_widths={
            "A": 16,
            "B": 30,
            "C": 8,
            "D": 8,
            "E": 32,
            "F": 28,
            "G": 18,
            "H": 14,
            "I": 16,
            "J": 16,
            "K": 16,
            "L": 16,
            "M": 22,
            "N": 22,
            "O": 18,
            "P": 18,
            "Q": 38,
            "S": 32,
            "T": 18,
            "U": 18,
            "V": 18,
            "W": 18,
            "X": 18,
            "Y": 18,
            "Z": 18,
            "AA": 18,
            "AB": 18,
            "AC": 24,
            "AD": 42,
        },
        skip_columns={"R"},
    )

    # Preservar freeze panes / sheet view definidos pelo template.
    ws_resultados.print_area = f"A1:AD{last_result_row}"

    output_path.parent.mkdir(parents=True, exist_ok=True)

    # =========================================================
    # GRAVAR E PROTEGER
    # =========================================================
    if not protect_output:
        wb.save(output_path)
        wb.close()
        validate_saved_workbook(output_path)
        return output_path

    report_password = build_report_password(week_start)

    temp_file = tempfile.NamedTemporaryFile(
        prefix=f".{output_path.stem}_",
        suffix=".xlsx",
        dir=output_path.parent,
        delete=False,
    )
    temp_plain_path = Path(temp_file.name)
    temp_file.close()

    try:
        wb.save(temp_plain_path)
        wb.close()
        validate_saved_workbook(temp_plain_path)

        encrypt_workbook(
            plain_path=temp_plain_path,
            encrypted_path=output_path,
            password=report_password,
        )

        validate_encrypted_workbook(output_path, report_password)

    finally:
        try:
            temp_plain_path.unlink(missing_ok=True)
        except Exception:
            pass

    return output_path
