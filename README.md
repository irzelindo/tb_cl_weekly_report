# Relatório semanal TB — multilaboratório

Implementação na branch **tb**. Laboratórios: **Carmelo, Nampula, Maputo e Beira**.

## Template oficial
Colocar o template aprovado em `templates/TB_Weekly_Report_Template.xlsx`.

Folhas: **Resumo, Comparação, TRL, Pendências, Resistência, Resultados, Desagregação Sexo**.

## Configuração
Instalar `pip install -r requirements.txt` e copiar `.env.example` para `.env`.
Configurar `REPORT_PASSWORD=tuberculose2026` como exemplo e **alterar a senha somente no `.env` local**.
A variável opcional `REPORT_SHEET_PASSWORD` protege a folha nominal contra edição.

## Executar
```bash
python src/generate_report.py --week-start 2026-09-13
```

São executadas cinco queries SQL por laboratório, todas com parâmetros `(week_start, lab)`.
O ficheiro final é encriptado para exigir senha na abertura; a folha Resultados recebe protecção contra edição.

## Regras e ressalvas
- **TRL**, não TAT: dias inteiros, arredondados apenas na apresentação. A classificação do cumprimento considera a média exacta.
- MGIT 44 dias; LJ 60 dias; LPA 7 dias; TSA 30 dias.
- Backlog à data de corte: amostras registadas até ao fim da semana, ainda sem resultado conhecido nessa data. Datas históricas podem estar incompletas.
- Pendências: 0–7, 8–44, **45–60**, >60 dias. Evita dupla contagem do dia 44.
- Rejeitadas: **em branco**, até validação de um código inequívoco no sistema.
- Atribuição do laboratório usa `TestingFacilityName`, senão `LIMSFacilityName`. As variantes reais devem ser validadas na base.
- Os campos por método MGIT/LJ podem não cobrir registos pendentes que ainda não identificam qual método será executado.
- Os totais somados de amostras por laboratório não representam automaticamente pacientes únicos.
- Nunca versionar `.env` ou relatórios nominais. Atenção: há um relatório já versionado na pasta reports; validar a confidencialidade.

## Validação
Executar primeiro os testes de SQL no servidor e comparar os quatro laboratórios, especialmente os registos pendentes e as designações das unidades.
Não considerar as métricas rejeitadas ou a desagregação por método como validadas até confirmação laboratorial.