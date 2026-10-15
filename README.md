# TB Weekly Report — Template Compacto Aprovado

## Ficheiros
- `templates/TB_Weekly_Report_Template.xlsx`
- `src/excel_writer.py`
- `src/generate_report.py`

## Mapeamento principal
- Resumo: visão geral, RR/MR/Pré-XR/XR, cultura MGIT/LJ, testes complementares, TAT resumido, resistência resumida.
- TAT: detalhe completo; inclui também `TSA 1 e 2` na linha 9.
- Resistência: detalhe por método.
- Resultados: line list em A:AC (29 colunas).

## Larguras
O Python **não executa AutoFit**. As larguras aprovadas no template são preservadas.

## Senha
A password continua rotativa:
`week_start + 7 dias`, por exemplo:
`13/09/2026` -> `20desetembrode2026`.

## Teste técnico sem senha
```bash
python src/generate_report.py --week-start 2026-09-13 --no-password
```

## Produção com senha
```bash
python src/generate_report.py --week-start 2026-09-13
```

## Nota
Na secção "Testes Complementares":
- LPA1 Total usa `LPA1_Total` se a Query 1 o devolver; caso contrário usa CMTB + MNT.
- TSA 1 / TSA 2 usa `TSA_Total` se disponível; caso contrário usa TSA1 + TSA2.
