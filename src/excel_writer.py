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
from openpyxl.styles import Alignment, Border, Font, Side


# =============================================================
# CONFIGURAÇÃO DO TEMPLATE COMPACTO APROVADO
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
    "LPA2 - Drogas": 17,                           # Q
    "LPA CM - Identificação": 18,                  # R
    "TSA1 - Streptomicina": 19,                    # S
    "TSA1 - Isoniazida": 20,                       # T
    "TSA1 - Rifampicina": 21,                      # U
    "TSA1 - Etambutol": 22,                        # V
    "TSA2 - Bedaquilina": 23,                      # W
    "TSA2 - Linezolida": 24,                       # X
    "TSA2 - Moxifloxacina": 25,                    # Y
    "TSA2 - Delamanida": 26,                       # Z
    "TSA2 - Clofazimina": 27,                      # AA
    "Classificação da Resistência": 28,            # AB
    "Link do Laudo": 29,                           # AC
}

RESULTADOS_DATE_FIELDS = {
    "Data de Colheita",
    "Data de Registo no Lab",
    "Data de Processamento",
    "Data de Validação",
}

THIN_GRAY = Side(style="thin", color="BFBFBF")


# =============================================================
# FUNÇÕES AUXILIARES
# =============================================================

def copy_cell_style(source, target):
    """Copia formatação sem escrever em células secundárias de merges."""
    if isinstance(source, MergedCell) or isinstance(target, MergedCell):
        return

    if source.has_style:
        target.font = copy(source.font)
        target.fill = copy(source.fill)
        target.border = copy(source.border)
        target.alignment = copy(source.alignment)
        target.number_format = source.number_format
        target.protection = copy(source.protection)


def copy_row_style(ws, source_row, target_row, start_col, end_col):
    """Copia a formatação de uma linha-modelo."""
    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height

    for col in range(start_col, end_col + 1):
        source = ws.cell(source_row, col)
        target = ws.cell(target_row, col)

        if isinstance(source, MergedCell) or isinstance(target, MergedCell):
            continue

        copy_cell_style(source, target)


def style_result_data_row(ws, row_num):
    """
    Garante formatação homogénea das linhas dinâmicas da folha Resultados.
    Não altera larguras das colunas do template.
    """
    for col in range(1, 30):
        cell = ws.cell(row_num, col)

        if isinstance(cell, MergedCell):
            continue

        font = copy(cell.font)
        font.name = "Cambria"
        font.sz = 10
        cell.font = font

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )

        cell.border = Border(
            left=THIN_GRAY,
            right=THIN_GRAY,
            top=THIN_GRAY,
            bottom=THIN_GRAY,
        )

    # Campos textuais alinhados à esquerda.
    for col in (2, 5, 6, 7, 17, 18, 28, 29):
        ws.cell(row_num, col).alignment = Alignment(
            horizontal="left",
            vertical="center",
            wrap_text=True,
        )


def set_hyperlink(cell, url):
    """Cria hyperlink sem aplicar o NamedStyle Hyperlink."""
    cell.hyperlink = url

    font = copy(cell.font)
    font.name = "Cambria"
    font.sz = 10
    font.color = "0563C1"
    font.underline = "single"
    cell.font = font


def set_value(cell, value, number_format=None, dash_if_none=False):
    """
    Escreve um valor preservando o estilo existente do template.

    dash_if_none=True:
        mostra "—" para não aplicável/sem dados.
    """
    if isinstance(cell, MergedCell):
        return

    if value is None:
        cell.value = "—" if dash_if_none else None
        if dash_if_none:
            cell.number_format = "General"
        return

    cell.value = value

    if number_format:
        cell.number_format = number_format


def safe_sum(*values):
    """Soma apenas valores numéricos; devolve None se todos forem None."""
    numeric = [
        value
        for value in values
        if isinstance(value, (int, float))
    ]

    if not numeric:
        return None

    return sum(numeric)


# =============================================================
# METADADOS — RESUMO
# =============================================================

def write_report_metadata(ws, week_start, week_end):
    """
    Layout aprovado:

      A4  Período do relatório:
      B4  valor
      C4  Nº do relatório:
      D4  valor

      A5  Data de emissão:
      B5  valor
      C5  Fonte dos dados:
      D5  valor já existente no template
    """

    ws["B4"] = (
        f"{week_start:%d/%m/%Y} "
        f"a "
        f"{week_end:%d/%m/%Y}"
    )

    report_number = week_start.strftime("%d%m%Y")
    ws["D4"] = report_number
    ws["D4"].number_format = "@"

    ws["B5"] = datetime.now().strftime("%d/%m/%Y")

    return report_number


# =============================================================
# QUERY 1 — RESUMO
# =============================================================

def write_indicadores_resumo(ws, rows):
    """
    Distribui as 5 linhas da Query 1 pelas secções compactas:

      Visão geral:
        linha 10 = Total de amostras processadas
        linha 11 = Total de amostras contaminadas

      Indicadores prioritários:
        linha 15 = RR/MR
        linha 16 = Pré-XR
        linha 17 = XR

      Resultado final da cultura:
        linha 21 = MGIT
        linha 22 = LJ

      Testes complementares:
        linha 26 = LPA1
        linha 27 = LPA2
        linha 28 = LPA CM
        linha 29 = TSA1 / TSA2
    """

    by_order = {
        row["Ordem"]: row
        for row in rows
    }

    # ---------------------------------------------------------
    # 1. VISÃO GERAL
    # ---------------------------------------------------------
    for order, excel_row in ((1, 10), (2, 11)):
        record = by_order.get(order)

        if not record:
            for col in range(2, 6):
                set_value(ws.cell(excel_row, col), None, dash_if_none=True)
            continue

        set_value(ws.cell(excel_row, 2), record.get("Homens"), "0")
        set_value(ws.cell(excel_row, 3), record.get("Mulheres"), "0")
        set_value(ws.cell(excel_row, 4), record.get("NaoEspecificado"), "0")
        set_value(ws.cell(excel_row, 5), record.get("Total"), "0")

    # ---------------------------------------------------------
    # 2. INDICADORES PRIORITÁRIOS
    # ---------------------------------------------------------
    for order, excel_row in ((3, 15), (4, 16), (5, 17)):
        record = by_order.get(order)

        if not record:
            for col in range(2, 6):
                set_value(ws.cell(excel_row, col), None, dash_if_none=True)
            continue

        set_value(ws.cell(excel_row, 2), record.get("Homens"), "0")
        set_value(ws.cell(excel_row, 3), record.get("Mulheres"), "0")
        set_value(ws.cell(excel_row, 4), record.get("NaoEspecificado"), "0")
        set_value(ws.cell(excel_row, 5), record.get("Total"), "0")

    # ---------------------------------------------------------
    # 3. RESULTADO FINAL DA CULTURA
    # ---------------------------------------------------------
    base = by_order.get(1)

    if base:
        # MGIT
        set_value(ws["B21"], base.get("Liquido_MGIT"), "0")
        set_value(ws["C21"], base.get("MGIT_CMTB"), "0")
        set_value(ws["D21"], base.get("MGIT_MNT"), "0")
        set_value(ws["E21"], base.get("MGIT_Negativo"), "0")

        # LJ
        set_value(ws["B22"], base.get("Solido_LJ"), "0")
        set_value(ws["C22"], base.get("LJ_CMTB"), "0")
        set_value(ws["D22"], base.get("LJ_MNT"), "0")
        set_value(ws["E22"], base.get("LJ_Negativo"), "0")

        # -----------------------------------------------------
        # 4. TESTES COMPLEMENTARES
        # -----------------------------------------------------
        # LPA1: categorias disponíveis na Query 1.
        lpa1_cmtb = base.get("LPA1_CMTB")
        lpa1_mnt = base.get("LPA1_MNT")

        set_value(ws["B26"], lpa1_cmtb, "0")
        set_value(ws["C26"], lpa1_mnt, "0")

        # Se a Query 1 futuramente devolver LPA1_Total, usamos-o.
        # Caso contrário, o total classificado é CMTB + MNT.
        lpa1_total = base.get("LPA1_Total")
        if lpa1_total is None:
            lpa1_total = safe_sum(lpa1_cmtb, lpa1_mnt)
        set_value(ws["D26"], lpa1_total, "0")

        # LPA2: CMTB/MNT não são o indicador usado nesta tabela.
        set_value(ws["B27"], None, dash_if_none=True)
        set_value(ws["C27"], None, dash_if_none=True)
        set_value(ws["D27"], base.get("LPA2"), "0")

        # LPA CM: identificação de espécie.
        set_value(ws["B28"], None, dash_if_none=True)
        set_value(ws["C28"], None, dash_if_none=True)
        set_value(ws["D28"], base.get("LPA_CM"), "0")

        # TSA 1 / TSA 2: total de determinações agregadas.
        set_value(ws["B29"], None, dash_if_none=True)
        set_value(ws["C29"], None, dash_if_none=True)

        tsa_total = base.get("TSA_Total")
        if tsa_total is None:
            tsa_total = safe_sum(
                base.get("TSA1"),
                base.get("TSA2"),
            )
        set_value(ws["D29"], tsa_total, "0")

    else:
        for cell_ref in (
            "B21", "C21", "D21", "E21",
            "B22", "C22", "D22", "E22",
            "B26", "C26", "D26",
            "B27", "C27", "D27",
            "B28", "C28", "D28",
            "B29", "C29", "D29",
        ):
            set_value(ws[cell_ref], None, dash_if_none=True)


# =============================================================
# QUERY 2 — TAT
# =============================================================

def _write_tat_record_summary(ws, excel_row, record):
    """Escreve uma linha TAT no resumo compacto."""
    if not record:
        for col in range(2, 9):
            set_value(ws.cell(excel_row, col), None, dash_if_none=True)
        return

    set_value(ws.cell(excel_row, 2), record.get("TotalRegistos"), "0")
    set_value(
        ws.cell(excel_row, 3),
        record.get("Media do tempo avaliado para o Target"),
        "0.00",
    )
    set_value(ws.cell(excel_row, 4), record.get("Target"))
    set_value(ws.cell(excel_row, 5), record.get("Registos fora do prazo"), "0")

    percentage = record.get("Percentagem fora do prazo")
    if percentage is None:
        set_value(ws.cell(excel_row, 6), None, dash_if_none=True)
    else:
        set_value(
            ws.cell(excel_row, 6),
            float(percentage) / 100.0,
            "0.00%",
        )

    set_value(
        ws.cell(excel_row, 7),
        record.get("Cumprimento do Target pela Media"),
        dash_if_none=True,
    )
    set_value(
        ws.cell(excel_row, 8),
        record.get("Alerta de Atrasos"),
        dash_if_none=True,
    )


def write_tat_resumo(ws, rows):
    """
    O resumo mantém apenas as 4 linhas mais importantes para leitura rápida.

      33 MGIT
      34 LJ
      35 LPA - Expectoração
      36 LPA - Cultura positiva

    A linha TSA continua disponível integralmente na folha TAT.
    """

    by_order = {
        row["Ordem"]: row
        for row in rows
    }

    for order, excel_row in (
        (1, 33),
        (2, 34),
        (3, 35),
        (4, 36),
    ):
        _write_tat_record_summary(
            ws,
            excel_row,
            by_order.get(order),
        )


def write_tat_detalhe(ws, rows):
    """
    Preenche a folha TAT.

    Mapeamento:
      A Exame
      B Total Registos
      C Com TAT válido
      D Sem TAT válido
      E Colheita→Registo
      F Registo→Processamento
      G Processamento→Validação
      H Total Colheita→Validação
      I Média p/ Target
      J Maior TAT
      K Nº Fora do Prazo
      L % Fora do Prazo
      M espaço
      N Target
      O Cumprimento
      P Alerta

    A Query 2 devolve também TSA (Ordem 5), gravado na linha 9.
    """

    by_order = {
        row["Ordem"]: row
        for row in rows
    }

    # Garantir estilo para a linha TSA (linha 9).
    copy_row_style(
        ws,
        source_row=8,
        target_row=9,
        start_col=1,
        end_col=16,
    )

    rows_map = {
        1: 5,
        2: 6,
        3: 7,
        4: 8,
        5: 9,
    }

    for order, excel_row in rows_map.items():
        record = by_order.get(order)

        if not record:
            for col in list(range(2, 13)) + list(range(14, 17)):
                set_value(ws.cell(excel_row, col), None, dash_if_none=True)
            continue

        set_value(ws.cell(excel_row, 1), record.get("Exame"), dash_if_none=True)
        set_value(ws.cell(excel_row, 2), record.get("TotalRegistos"), "0")
        set_value(ws.cell(excel_row, 3), record.get("RegistosComTempoValido"), "0")
        set_value(
            ws.cell(excel_row, 4),
            record.get("Registos sem tempo valido"),
            "0",
        )
        set_value(
            ws.cell(excel_row, 5),
            record.get("Media Colheita ao Registo"),
            "0.00",
        )
        set_value(
            ws.cell(excel_row, 6),
            record.get("Media Registo ao Processamento"),
            "0.00",
        )
        set_value(
            ws.cell(excel_row, 7),
            record.get("Media Processamento a Validacao"),
            "0.00",
        )
        set_value(
            ws.cell(excel_row, 8),
            record.get("Media Total Colheita a Validacao"),
            "0.00",
        )
        set_value(
            ws.cell(excel_row, 9),
            record.get("Media do tempo avaliado para o Target"),
            "0.00",
        )
        set_value(
            ws.cell(excel_row, 10),
            record.get("Maior tempo para o Target"),
            "0",
        )
        set_value(
            ws.cell(excel_row, 11),
            record.get("Registos fora do prazo"),
            "0",
        )

        percentage = record.get("Percentagem fora do prazo")
        if percentage is None:
            set_value(ws.cell(excel_row, 12), None, dash_if_none=True)
        else:
            set_value(
                ws.cell(excel_row, 12),
                float(percentage) / 100.0,
                "0.00%",
            )

        set_value(ws.cell(excel_row, 14), record.get("Target"), dash_if_none=True)
        set_value(
            ws.cell(excel_row, 15),
            record.get("Cumprimento do Target pela Media"),
            dash_if_none=True,
        )
        set_value(
            ws.cell(excel_row, 16),
            record.get("Alerta de Atrasos"),
            dash_if_none=True,
        )

    # Ajuste visual mínimo da linha TSA.
    ws["A9"] = by_order.get(5, {}).get("Exame") if by_order.get(5) else "TSA 1 e 2"


# =============================================================
# QUERY 3 — RESISTÊNCIA
# =============================================================

def _write_resistance_row(
    ws,
    excel_row,
    record,
    detailed=False,
):
    if not record:
        max_col = 9 if detailed else 5
        for col in range(2, max_col + 1):
            set_value(ws.cell(excel_row, col), None, dash_if_none=True)
        return

    if detailed:
        fields = {
            1: "Medicamento",
            2: "LPA1",
            3: "LPA2",
            4: "TSA1",
            5: "TSA2",
            6: "TotalResistente",
            7: "NumeroPacientes",
            8: "NumeroTestes",
        }

        for col, field in fields.items():
            set_value(
                ws.cell(excel_row, col),
                record.get(field),
                "0" if col > 1 else None,
                dash_if_none=True,
            )

        percentage = record.get("PercentagemResistencia")
        if percentage is None:
            set_value(ws.cell(excel_row, 9), None, dash_if_none=True)
        else:
            set_value(
                ws.cell(excel_row, 9),
                float(percentage) / 100.0,
                "0.00%",
            )

    else:
        # Resumo: A medicamento já está no template.
        set_value(ws.cell(excel_row, 2), record.get("TotalResistente"), "0", True)
        set_value(ws.cell(excel_row, 3), record.get("NumeroPacientes"), "0", True)
        set_value(ws.cell(excel_row, 4), record.get("NumeroTestes"), "0", True)

        percentage = record.get("PercentagemResistencia")
        if percentage is None:
            set_value(ws.cell(excel_row, 5), None, dash_if_none=True)
        else:
            set_value(
                ws.cell(excel_row, 5),
                float(percentage) / 100.0,
                "0.00%",
            )


def write_resistencia_resumo(ws, rows):
    by_order = {
        row["Ordem"]: row
        for row in rows
    }

    for order in range(1, 11):
        excel_row = 39 + order  # 40..49
        _write_resistance_row(
            ws,
            excel_row,
            by_order.get(order),
            detailed=False,
        )


def write_resistencia_detalhe(ws, rows):
    by_order = {
        row["Ordem"]: row
        for row in rows
    }

    for order in range(1, 11):
        excel_row = 4 + order  # 5..14
        _write_resistance_row(
            ws,
            excel_row,
            by_order.get(order),
            detailed=True,
        )


# =============================================================
# QUERY 4 — RESULTADOS
# =============================================================

def write_resultados(ws, rows):
    """
    Folha Resultados do template compacto:
      cabeçalho = linha 2
      dados     = linha 3+
      colunas   = A:AC (29 colunas)
    """

    DATA_START_ROW = 3
    last_data_row = (
        DATA_START_ROW + len(rows) - 1
        if rows
        else DATA_START_ROW
    )

    existing_last_row = max(ws.max_row, DATA_START_ROW)

    # Limpar valores antigos.
    clear_until = max(
        existing_last_row,
        last_data_row,
    )

    for row_num in range(
        DATA_START_ROW,
        clear_until + 1,
    ):
        for col_num in range(1, 30):
            cell = ws.cell(row_num, col_num)

            if isinstance(cell, MergedCell):
                continue

            cell.value = None
            cell.hyperlink = None

    # Escrever dados.
    for index, record in enumerate(rows):
        excel_row = DATA_START_ROW + index

        style_result_data_row(
            ws,
            excel_row,
        )

        for field, column in RESULTADOS_MAPPING.items():
            cell = ws.cell(
                excel_row,
                column,
            )

            value = record.get(field)
            cell.value = value

            if (
                field in RESULTADOS_DATE_FIELDS
                and value is not None
            ):
                cell.number_format = "dd/mm/yyyy"

            if (
                field == "Link do Laudo"
                and value
            ):
                set_hyperlink(
                    cell,
                    value,
                )

    if rows:
        ws.auto_filter.ref = (
            f"A2:AC{last_data_row}"
        )
    else:
        ws.auto_filter.ref = "A2:AC3"

    ws.print_area = (
        f"A1:AC{last_data_row}"
    )

    return last_data_row


# =============================================================
# PASSWORD ROTATIVA
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
    match = re.search(
        r"(\d{4})-(\d{2})-(\d{2})",
        Path(template_path).name,
    )

    if not match:
        return None

    try:
        start = date(
            int(match.group(1)),
            int(match.group(2)),
            int(match.group(3)),
        )
    except ValueError:
        return None

    return build_report_password(start)


def load_template_workbook(
    template_path,
    week_start,
    template_password=None,
):
    """
    Abre template simples ou encriptado.

    Para produção recomenda-se manter o template mestre sem senha
    e proteger apenas o relatório final.
    """

    template_path = Path(template_path)

    if zipfile.is_zipfile(template_path):
        return load_workbook(template_path)

    try:
        import msoffcrypto
    except ImportError as exc:
        raise RuntimeError(
            "O template parece estar encriptado. "
            "Instale 'msoffcrypto-tool' ou use um template sem senha."
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
                office_file.load_key(
                    password=password,
                    verify_password=True,
                )
                office_file.decrypt(
                    decrypted,
                    verify_integrity=True,
                )

            decrypted.seek(0)
            return load_workbook(decrypted)

        except Exception as exc:
            last_error = exc

    raise RuntimeError(
        "Não foi possível abrir o template encriptado."
    ) from last_error


def encrypt_workbook(
    plain_path,
    encrypted_path,
    password,
):
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
            office_file.encrypt(
                password,
                destination,
            )


def validate_saved_workbook(path):
    """Valida o XLSX antes da encriptação."""
    path = Path(path)

    with zipfile.ZipFile(path, "r") as archive:
        bad_file = archive.testzip()

        if bad_file is not None:
            raise ValueError(
                f"XLSX corrompido: {bad_file}"
            )

    validation_wb = load_workbook(
        path,
        read_only=False,
        data_only=False,
    )
    validation_wb.close()


def validate_encrypted_workbook(
    path,
    password,
):
    """Valida o relatório final protegido sem criar cópia desencriptada em disco."""
    try:
        import msoffcrypto
    except ImportError as exc:
        raise RuntimeError(
            "Validação requer 'msoffcrypto-tool'."
        ) from exc

    decrypted = io.BytesIO()

    with Path(path).open("rb") as source:
        office_file = msoffcrypto.OfficeFile(
            source
        )

        if not office_file.is_encrypted():
            raise ValueError(
                "O relatório final não está encriptado."
            )

        office_file.load_key(
            password=password,
            verify_password=True,
        )

        office_file.decrypt(
            decrypted,
            verify_integrity=True,
        )

    decrypted.seek(0)

    if not zipfile.is_zipfile(decrypted):
        raise ValueError(
            "O conteúdo desencriptado não é um XLSX válido."
        )


# =============================================================
# GERAÇÃO FINAL
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
    """
    Gera o relatório semanal usando o template compacto aprovado.

    Importante:
      - preserva as larguras definidas manualmente no template;
      - não executa AutoFit;
      - preserva freeze panes / sheet views do template;
      - preenche as folhas Resumo, TAT, Resistência e Resultados.
    """

    template_path = Path(
        template_path
    )

    output_path = Path(
        output_path
    )

    if not template_path.exists():
        raise FileNotFoundError(
            f"Template não encontrado: {template_path}"
        )

    wb = load_template_workbook(
        template_path=template_path,
        week_start=week_start,
        template_password=template_password,
    )

    required_sheets = {
        "Resumo",
        "TAT",
        "Resistência",
        "Resultados",
    }

    missing = (
        required_sheets
        - set(wb.sheetnames)
    )

    if missing:
        raise KeyError(
            "Folhas em falta no template: "
            + ", ".join(sorted(missing))
        )

    ws_resumo = wb["Resumo"]
    ws_tat = wb["TAT"]
    ws_resistencia = wb["Resistência"]
    ws_resultados = wb["Resultados"]

    # ---------------------------------------------------------
    # Preenchimento
    # ---------------------------------------------------------
    write_report_metadata(
        ws_resumo,
        week_start,
        week_end,
    )

    write_indicadores_resumo(
        ws_resumo,
        indicadores,
    )

    write_tat_resumo(
        ws_resumo,
        tat,
    )

    write_resistencia_resumo(
        ws_resumo,
        resistencia,
    )

    write_tat_detalhe(
        ws_tat,
        tat,
    )

    write_resistencia_detalhe(
        ws_resistencia,
        resistencia,
    )

    write_resultados(
        ws_resultados,
        resultados,
    )

    # ---------------------------------------------------------
    # Não alterar larguras, freeze panes ou sheet views.
    # São controlados pelo template aprovado.
    # ---------------------------------------------------------

    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    # ---------------------------------------------------------
    # Sem password — útil para validação técnica.
    # ---------------------------------------------------------
    if not protect_output:
        wb.save(output_path)
        wb.close()

        validate_saved_workbook(
            output_path
        )

        return output_path

    # ---------------------------------------------------------
    # Com password rotativa.
    # ---------------------------------------------------------
    report_password = build_report_password(
        week_start
    )

    temp_file = tempfile.NamedTemporaryFile(
        prefix=f".{output_path.stem}_",
        suffix=".xlsx",
        dir=output_path.parent,
        delete=False,
    )

    temp_plain_path = Path(
        temp_file.name
    )

    temp_file.close()

    try:
        wb.save(
            temp_plain_path
        )

        wb.close()

        validate_saved_workbook(
            temp_plain_path
        )

        encrypt_workbook(
            plain_path=temp_plain_path,
            encrypted_path=output_path,
            password=report_password,
        )

        validate_encrypted_workbook(
            output_path,
            report_password,
        )

    finally:
        try:
            temp_plain_path.unlink(
                missing_ok=True
            )
        except Exception:
            pass

    return output_path
