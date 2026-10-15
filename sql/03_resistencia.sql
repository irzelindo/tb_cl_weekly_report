SET NOCOUNT ON;

/* ============================================================
   PARÂMETROS RECEBIDOS DO PYTHON
   ============================================================ */
DECLARE @WeekStart DATE = ?;
DECLARE @WeekEnd   DATE = DATEADD(DAY, 7, @WeekStart);


;WITH Base AS
(
    SELECT
          T.ID
        , T.RequestID
        , T.EncryptedPatientID
        , T.PatientUUID
        , T.PatientNationalID
        , T.PatientHealthcareNo

        , T.TSAStreptomycinResult
        , T.LPAIsoniazidResult
        , T.TSAIsoniazidResult
        , T.LPARifampicinResult
        , T.TSARifampicinResult
        , T.TSAEthambutolResult
        , T.TSA2BedaquilineResult
        , T.TSA2LinezolidResult
        , T.LPA2FluoroquinoloneResult
        , T.TSA2MoxifloxacinResult
        , T.TSA2ClofazimineResult
        , T.TSA2ProthionamideResult

        , CAST
          (
              CASE
                  WHEN NULLIF(LTRIM(RTRIM(T.EncryptedPatientID)), '') IS NOT NULL
                      THEN 'ENC:' + LTRIM(RTRIM(T.EncryptedPatientID))
                  WHEN NULLIF(LTRIM(RTRIM(T.PatientUUID)), '') IS NOT NULL
                      THEN 'UUID:' + LTRIM(RTRIM(T.PatientUUID))
                  WHEN NULLIF(LTRIM(RTRIM(T.PatientNationalID)), '') IS NOT NULL
                      THEN 'NID:' + LTRIM(RTRIM(T.PatientNationalID))
                  WHEN NULLIF(LTRIM(RTRIM(T.PatientHealthcareNo)), '') IS NOT NULL
                      THEN 'HC:' + LTRIM(RTRIM(T.PatientHealthcareNo))
                  ELSE 'REQ:' + T.RequestID
              END
              AS VARCHAR(150)
          ) AS PatientKey

    FROM dbo.TBCultura AS T
    WHERE T.HL7ResultStatusCode = 'F'
      AND T.AnalysisDateTime >= @WeekStart
      AND T.AnalysisDateTime <  @WeekEnd
),
TestesBrutos AS
(
    SELECT
          B.ID
        , B.RequestID
        , B.PatientKey
        , X.Medicamento
        , X.Metodo
        , LTRIM(RTRIM(X.Resultado)) AS ResultadoBruto
    FROM Base AS B
    CROSS APPLY
    (
        VALUES
            ('Streptomicina', 'TSA1', B.TSAStreptomycinResult),

            ('Isoniazida', 'LPA1', B.LPAIsoniazidResult),
            ('Isoniazida', 'TSA1', B.TSAIsoniazidResult),

            ('Rifampicina', 'LPA1', B.LPARifampicinResult),
            ('Rifampicina', 'TSA1', B.TSARifampicinResult),

            ('Etambutol', 'TSA1', B.TSAEthambutolResult),

            ('Bedaquilina', 'TSA2', B.TSA2BedaquilineResult),
            ('Linezolida', 'TSA2', B.TSA2LinezolidResult),

            /* LPA2 é resultado da classe das fluoroquinolonas. */
            ('Moxifloxacina', 'LPA2', B.LPA2FluoroquinoloneResult),
            ('Moxifloxacina', 'TSA2', B.TSA2MoxifloxacinResult),

            ('Clofazimina', 'TSA2', B.TSA2ClofazimineResult),
            ('Protionamida', 'TSA2', B.TSA2ProthionamideResult)
    ) AS X(Medicamento, Metodo, Resultado)
    WHERE NULLIF(LTRIM(RTRIM(X.Resultado)), '') IS NOT NULL
),
Testes AS
(
    /*
       Apenas resultados de susceptibilidade válidos contam como teste.
       Isto evita contar textos como "Não realizado" como determinação.
    */
    SELECT
          ID
        , RequestID
        , PatientKey
        , Medicamento
        , Metodo
        , CASE
              WHEN ResultadoBruto IN ('Resistant', 'Resistente')
                  THEN 'Resistente'
              WHEN ResultadoBruto IN ('Susceptible', 'Sensível', 'Sensivel')
                  THEN 'Sensível'
              WHEN ResultadoBruto IN ('Intermediate', 'Intermédio', 'Intermedio')
                  THEN 'Intermédio'
          END AS Resultado
    FROM TestesBrutos
    WHERE ResultadoBruto IN
          (
              'Resistant', 'Resistente',
              'Susceptible', 'Sensível', 'Sensivel',
              'Intermediate', 'Intermédio', 'Intermedio'
          )
),
Medicamentos AS
(
    /*
       Os indicadores de aplicabilidade permitem distinguir:
       - 0: método aplicável, mas sem resistentes;
       - NULL: método não se aplica àquele medicamento.
    */
    SELECT 1 AS Ordem, 'Streptomicina' AS Medicamento, 0 AS AplicaLPA1, 0 AS AplicaLPA2, 1 AS AplicaTSA1, 0 AS AplicaTSA2
    UNION ALL SELECT 2,  'Isoniazida',     1, 0, 1, 0
    UNION ALL SELECT 3,  'Rifampicina',    1, 0, 1, 0
    UNION ALL SELECT 4,  'Etambutol',      0, 0, 1, 0
    UNION ALL SELECT 5,  'Bedaquilina',    0, 0, 0, 1
    UNION ALL SELECT 6,  'Linezolida',     0, 0, 0, 1
    UNION ALL SELECT 7,  'Moxifloxacina',  0, 1, 0, 1
    UNION ALL SELECT 8,  'Delamanida',     0, 0, 0, 0
    UNION ALL SELECT 9,  'Clofazimina',    0, 0, 0, 1
    UNION ALL SELECT 10, 'Protionamida',   0, 0, 0, 1
)
SELECT
      @WeekStart AS DataInicio
    , DATEADD(DAY, -1, @WeekEnd) AS DataFim
    , M.Ordem
    , M.Medicamento

    /* Pacientes resistentes por método. */
    , CASE
          WHEN M.AplicaLPA1 = 0 THEN NULL
          ELSE COUNT(DISTINCT CASE WHEN T.Metodo = 'LPA1' AND T.Resultado = 'Resistente' THEN T.PatientKey END)
      END AS LPA1

    , CASE
          WHEN M.AplicaLPA2 = 0 THEN NULL
          ELSE COUNT(DISTINCT CASE WHEN T.Metodo = 'LPA2' AND T.Resultado = 'Resistente' THEN T.PatientKey END)
      END AS LPA2

    , CASE
          WHEN M.AplicaTSA1 = 0 THEN NULL
          ELSE COUNT(DISTINCT CASE WHEN T.Metodo = 'TSA1' AND T.Resultado = 'Resistente' THEN T.PatientKey END)
      END AS TSA1

    , CASE
          WHEN M.AplicaTSA2 = 0 THEN NULL
          ELSE COUNT(DISTINCT CASE WHEN T.Metodo = 'TSA2' AND T.Resultado = 'Resistente' THEN T.PatientKey END)
      END AS TSA2

    /* Pacientes únicos resistentes, independentemente do método. */
    , CASE
          WHEN M.Medicamento = 'Delamanida' THEN NULL
          ELSE COUNT(DISTINCT CASE WHEN T.Resultado = 'Resistente' THEN T.PatientKey END)
      END AS TotalResistente

    /* Pacientes únicos efectivamente testados. */
    , CASE
          WHEN M.Medicamento = 'Delamanida' THEN NULL
          ELSE COUNT(DISTINCT T.PatientKey)
      END AS NumeroPacientes

    /* Determinações válidas realizadas. */
    , CASE
          WHEN M.Medicamento = 'Delamanida' THEN NULL
          ELSE COUNT(T.Resultado)
      END AS NumeroTestes

    /* Pacientes resistentes / pacientes testados. */
    , CASE
          WHEN M.Medicamento = 'Delamanida' THEN NULL
          ELSE CAST
          (
              100.0
              * COUNT(DISTINCT CASE WHEN T.Resultado = 'Resistente' THEN T.PatientKey END)
              / NULLIF(COUNT(DISTINCT T.PatientKey), 0)
              AS DECIMAL(6,2)
          )
      END AS PercentagemResistencia

FROM Medicamentos AS M
LEFT JOIN Testes AS T
    ON T.Medicamento = M.Medicamento
GROUP BY
      M.Ordem
    , M.Medicamento
    , M.AplicaLPA1
    , M.AplicaLPA2
    , M.AplicaTSA1
    , M.AplicaTSA2
ORDER BY M.Ordem;
