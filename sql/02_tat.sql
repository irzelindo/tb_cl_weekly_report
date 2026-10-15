SET NOCOUNT ON;

/* ============================================================
   PARÂMETROS RECEBIDOS DO PYTHON
   ============================================================ */
DECLARE @WeekStart DATE = ?;
DECLARE @LabName NVARCHAR(20) = ?;
DECLARE @WeekEnd   DATE = DATEADD(DAY, 7, @WeekStart);

/* Targets operacionais.
   LJ = 60 dias conforme template validado pelo grupo. */
DECLARE @MGITTargetDias INT = 44;
DECLARE @LJTargetDias   INT = 60;
DECLARE @LPATargetDias  INT = 7;
DECLARE @TSATargetDias  INT = 30;


;WITH Base AS
(
    SELECT
          T.ID
        , T.RequestID
        , T.SpecimenDateTime
        , T.RegisteredDateTime
        , T.AnalysisDateTime
        , T.AuthorisedDateTime
        , T.LIMSSpecimenSourceDesc

        , T.LiquidCultureResult
        , T.SolidCultureResult
        , T.FinalMGITCultureResult
        , T.FinalMGITResultDate
        , T.FinalLJCultureResult
        , T.FinalLJResultDate

        , T.LPAComplexMTBResult
        , T.LPAIsoniazidResult
        , T.LPARifampicinResult
        , T.LPAResultDate
        , T.LPA2ComplexMTBResult
        , T.LPA2FluoroquinoloneResult
        , T.LPA2ResultDate

        , T.TSAEthambutolResult
        , T.TSAIsoniazidResult
        , T.TSARifampicinResult
        , T.TSAStreptomycinResult
        , T.TSAResultDate

        , T.TSA2AmikacinResult
        , T.TSA2BedaquilineResult
        , T.TSA2LinezolidResult
        , T.TSA2MoxifloxacinResult
        , T.TSA2ResultDate

        /* Datas de resultado em formato DATE. */
        , TRY_CONVERT(DATE, NULLIF(LTRIM(RTRIM(T.FinalMGITResultDate)), ''), 103) AS FinalMGITDate
        , TRY_CONVERT(DATE, NULLIF(LTRIM(RTRIM(T.FinalLJResultDate)), ''), 103) AS FinalLJDate
        , TRY_CONVERT(DATE, NULLIF(LTRIM(RTRIM(T.LPAResultDate)), ''), 103) AS LPA1Date
        , TRY_CONVERT(DATE, NULLIF(LTRIM(RTRIM(T.LPA2ResultDate)), ''), 103) AS LPA2Date
        , TRY_CONVERT(DATE, NULLIF(LTRIM(RTRIM(T.TSAResultDate)), ''), 103) AS TSA1Date
        , TRY_CONVERT(DATE, NULLIF(LTRIM(RTRIM(T.TSA2ResultDate)), ''), 103) AS TSA2Date

        /* Cultura CMTB positiva para determinar TAT após cultura positiva. */
        , CASE
              WHEN LTRIM(RTRIM(T.FinalMGITCultureResult)) IN
                   (
                       'Positive for MTB Complex',
                       'Positivo para Complexo MTB',
                       'Positivo para CMTB',
                       'C. Mycobacterium tuberculosis*'
                   )
                  THEN 1
              ELSE 0
          END AS MGITPositivo

        , CASE
              WHEN LTRIM(RTRIM(T.FinalLJCultureResult)) IN
                   (
                       'Positive for MTB Complex',
                       'Positivo para Complexo MTB',
                       'Positivo para CMTB',
                       'C. Mycobacterium tuberculosis*'
                   )
                  THEN 1
              ELSE 0
          END AS LJPositivo

        , CASE
              WHEN NULLIF(LTRIM(RTRIM(T.LiquidCultureResult)), '') IS NOT NULL
                OR NULLIF(LTRIM(RTRIM(T.FinalMGITCultureResult)), '') IS NOT NULL
                OR NULLIF(LTRIM(RTRIM(T.FinalMGITResultDate)), '') IS NOT NULL
                  THEN 1
              ELSE 0
          END AS HasMGIT

        , CASE
              WHEN NULLIF(LTRIM(RTRIM(T.SolidCultureResult)), '') IS NOT NULL
                OR NULLIF(LTRIM(RTRIM(T.FinalLJCultureResult)), '') IS NOT NULL
                OR NULLIF(LTRIM(RTRIM(T.FinalLJResultDate)), '') IS NOT NULL
                  THEN 1
              ELSE 0
          END AS HasLJ

        , CASE
              WHEN NULLIF(LTRIM(RTRIM(T.LPAComplexMTBResult)), '') IS NOT NULL
                OR NULLIF(LTRIM(RTRIM(T.LPAIsoniazidResult)), '') IS NOT NULL
                OR NULLIF(LTRIM(RTRIM(T.LPARifampicinResult)), '') IS NOT NULL
                OR NULLIF(LTRIM(RTRIM(T.LPA2ComplexMTBResult)), '') IS NOT NULL
                OR NULLIF(LTRIM(RTRIM(T.LPA2FluoroquinoloneResult)), '') IS NOT NULL
                  THEN 1
              ELSE 0
          END AS HasLPA

        , CASE
              WHEN NULLIF(LTRIM(RTRIM(T.TSAEthambutolResult)), '') IS NOT NULL
                OR NULLIF(LTRIM(RTRIM(T.TSAIsoniazidResult)), '') IS NOT NULL
                OR NULLIF(LTRIM(RTRIM(T.TSARifampicinResult)), '') IS NOT NULL
                OR NULLIF(LTRIM(RTRIM(T.TSAStreptomycinResult)), '') IS NOT NULL
                OR NULLIF(LTRIM(RTRIM(T.TSA2AmikacinResult)), '') IS NOT NULL
                OR NULLIF(LTRIM(RTRIM(T.TSA2BedaquilineResult)), '') IS NOT NULL
                OR NULLIF(LTRIM(RTRIM(T.TSA2LinezolidResult)), '') IS NOT NULL
                OR NULLIF(LTRIM(RTRIM(T.TSA2MoxifloxacinResult)), '') IS NOT NULL
                  THEN 1
              ELSE 0
          END AS HasTSA

        , CASE
              WHEN UPPER(ISNULL(T.LIMSSpecimenSourceDesc, '')) LIKE '%EXPECTOR%'
                OR UPPER(ISNULL(T.LIMSSpecimenSourceDesc, '')) LIKE '%ESCARR%'
                OR UPPER(ISNULL(T.LIMSSpecimenSourceDesc, '')) LIKE '%SPUT%'
                  THEN 1
              ELSE 0
          END AS IsSputum

    FROM dbo.TBCultura AS T
    CROSS APPLY (SELECT UPPER(CONVERT(NVARCHAR(255), COALESCE(
        NULLIF(LTRIM(RTRIM(T.TestingFacilityName)), ''),
        NULLIF(LTRIM(RTRIM(T.LIMSFacilityName)), ''), N''
    ))) COLLATE Latin1_General_CI_AI AS NomeBruto) AS LabRaw
    CROSS APPLY (SELECT CASE
        WHEN LabRaw.NomeBruto LIKE N'%CARMELO%' THEN N'Carmelo'
        WHEN LabRaw.NomeBruto LIKE N'%NAMPULA%' THEN N'Nampula'
        WHEN LabRaw.NomeBruto LIKE N'%PTB%INS%'
          OR LabRaw.NomeBruto IN (N'TB LABORATORIO', N'TB LABORATÓRIO')
            THEN N'Maputo'
        WHEN LabRaw.NomeBruto LIKE N'%BEIRA%' THEN N'Beira'
        ELSE NULL
    END AS Nome) AS Lab

    WHERE Lab.Nome = @LabName
      AND T.HL7ResultStatusCode = 'F'
      AND T.AnalysisDateTime >= @WeekStart
      AND T.AnalysisDateTime <  @WeekEnd
),
Datas AS
(
    SELECT
          B.*

        /* Primeira data de cultura CMTB positiva. */
        , CASE
              WHEN B.MGITPositivo = 1 AND B.LJPositivo = 1 THEN
                   CASE
                       WHEN B.FinalMGITDate IS NULL THEN B.FinalLJDate
                       WHEN B.FinalLJDate IS NULL THEN B.FinalMGITDate
                       WHEN B.FinalMGITDate <= B.FinalLJDate THEN B.FinalMGITDate
                       ELSE B.FinalLJDate
                   END
              WHEN B.MGITPositivo = 1 THEN B.FinalMGITDate
              WHEN B.LJPositivo = 1 THEN B.FinalLJDate
              ELSE NULL
          END AS DataCulturaPositiva

        /* Datas específicas de validação/resultado por método. */
        , COALESCE(B.FinalMGITDate, CAST(B.AuthorisedDateTime AS DATE)) AS MGITValidationDate
        , COALESCE(B.FinalLJDate, CAST(B.AuthorisedDateTime AS DATE)) AS LJValidationDate
        , COALESCE(B.LPA1Date, B.LPA2Date, CAST(B.AuthorisedDateTime AS DATE)) AS LPAValidationDate

        , CASE
              WHEN B.TSA1Date IS NULL THEN COALESCE(B.TSA2Date, CAST(B.AuthorisedDateTime AS DATE))
              WHEN B.TSA2Date IS NULL THEN B.TSA1Date
              WHEN B.TSA1Date >= B.TSA2Date THEN B.TSA1Date
              ELSE B.TSA2Date
          END AS TSAValidationDate

    FROM Base AS B
),
Exames AS
(
    /* ========================================================
       1. CULTURA LÍQUIDA (MGIT)
       Target: 44 dias do registo ao resultado final.
       ======================================================== */
    SELECT
          ID
        , 1 AS Ordem
        , 'Cultura Líquida (MGIT)' AS Exame
        , SpecimenDateTime
        , RegisteredDateTime
        , AnalysisDateTime
        , MGITValidationDate AS ValidationDate
        , CAST(RegisteredDateTime AS DATE) AS TargetStartDate
        , @MGITTargetDias AS TargetDias
    FROM Datas
    WHERE HasMGIT = 1

    UNION ALL

    /* ========================================================
       2. CULTURA SÓLIDA (LJ)
       Target operacional configurável: 60 dias.
       ======================================================== */
    SELECT
          ID
        , 2
        , 'Cultura Sólida (LJ)'
        , SpecimenDateTime
        , RegisteredDateTime
        , AnalysisDateTime
        , LJValidationDate
        , CAST(RegisteredDateTime AS DATE)
        , @LJTargetDias
    FROM Datas
    WHERE HasLJ = 1

    UNION ALL

    /* ========================================================
       3. LPA DIRECTO EM EXPECTORAÇÃO

       Esta população é mutuamente exclusiva da linha
       "LPA - Cultura Positiva": o LPA é classificado como
       directo quando não existe cultura CMTB positiva anterior
       ao resultado do LPA.
       ======================================================== */
    SELECT
          ID
        , 3
        , 'LPA - Amostras de Expectoração'
        , SpecimenDateTime
        , RegisteredDateTime
        , AnalysisDateTime
        , LPAValidationDate
        , CAST(SpecimenDateTime AS DATE)
        , @LPATargetDias
    FROM Datas
    WHERE HasLPA = 1
      AND IsSputum = 1
      AND
      (
          DataCulturaPositiva IS NULL
          OR (LPAValidationDate IS NOT NULL AND LPAValidationDate < DataCulturaPositiva)
      )

    UNION ALL

    /* ========================================================
       4. LPA APÓS CULTURA CMTB POSITIVA
       ======================================================== */
    SELECT
          ID
        , 4
        , 'LPA - Cultura Positiva'
        , SpecimenDateTime
        , RegisteredDateTime
        , AnalysisDateTime
        , LPAValidationDate
        , DataCulturaPositiva
        , @LPATargetDias
    FROM Datas
    WHERE HasLPA = 1
      AND DataCulturaPositiva IS NOT NULL
      AND
      (
          LPAValidationDate IS NULL
          OR LPAValidationDate >= DataCulturaPositiva
      )

    UNION ALL

    /* ========================================================
       5. TSA 1 E 2 APÓS CULTURA POSITIVA
       ======================================================== */
    SELECT
          ID
        , 5
        , 'TSA 1 e 2'
        , SpecimenDateTime
        , RegisteredDateTime
        , AnalysisDateTime
        , TSAValidationDate
        , DataCulturaPositiva
        , @TSATargetDias
    FROM Datas
    WHERE HasTSA = 1
      AND DataCulturaPositiva IS NOT NULL
),
Tempos AS
(
    SELECT
          E.*

        , CASE
              WHEN E.SpecimenDateTime IS NOT NULL
               AND E.RegisteredDateTime IS NOT NULL
               AND E.RegisteredDateTime >= E.SpecimenDateTime
                  THEN DATEDIFF(DAY, E.SpecimenDateTime, E.RegisteredDateTime)
          END AS ColheitaRegisto

        , CASE
              WHEN E.RegisteredDateTime IS NOT NULL
               AND E.AnalysisDateTime IS NOT NULL
               AND E.AnalysisDateTime >= E.RegisteredDateTime
                  THEN DATEDIFF(DAY, E.RegisteredDateTime, E.AnalysisDateTime)
          END AS RegistoProcessamento

        , CASE
              WHEN E.AnalysisDateTime IS NOT NULL
               AND E.ValidationDate IS NOT NULL
               AND E.ValidationDate >= CAST(E.AnalysisDateTime AS DATE)
                  THEN DATEDIFF(DAY, CAST(E.AnalysisDateTime AS DATE), E.ValidationDate)
          END AS ProcessamentoValidacao

        , CASE
              WHEN E.SpecimenDateTime IS NOT NULL
               AND E.ValidationDate IS NOT NULL
               AND E.ValidationDate >= CAST(E.SpecimenDateTime AS DATE)
                  THEN DATEDIFF(DAY, CAST(E.SpecimenDateTime AS DATE), E.ValidationDate)
          END AS ColheitaValidacao

        , CASE
              WHEN E.TargetStartDate IS NOT NULL
               AND E.ValidationDate IS NOT NULL
               AND E.ValidationDate >= E.TargetStartDate
                  THEN DATEDIFF(DAY, E.TargetStartDate, E.ValidationDate)
          END AS TempoTarget

    FROM Exames AS E
),
Estatisticas AS
(
    SELECT
          Ordem
        , Exame
        , TargetDias
        , COUNT(*) AS TotalRegistos
        , COUNT(TempoTarget) AS RegistosComTempoValido

        , AVG(CAST(ColheitaRegisto AS DECIMAL(10,2))) AS MediaColheitaRegisto
        , AVG(CAST(RegistoProcessamento AS DECIMAL(10,2))) AS MediaRegistoProcessamento
        , AVG(CAST(ProcessamentoValidacao AS DECIMAL(10,2))) AS MediaProcessamentoValidacao
        , AVG(CAST(ColheitaValidacao AS DECIMAL(10,2))) AS MediaTotal
        , AVG(CAST(TempoTarget AS DECIMAL(10,2))) AS MediaTarget

        , MAX(TempoTarget) AS MaiorTempoTarget
        , SUM(CASE WHEN TempoTarget > TargetDias THEN 1 ELSE 0 END) AS RegistosForaPrazo

    FROM Tempos
    GROUP BY
          Ordem
        , Exame
        , TargetDias
)
SELECT
      @WeekStart AS DataInicio
    , DATEADD(DAY, -1, @WeekEnd) AS DataFim
    , Ordem
    , Exame
    , TotalRegistos
    , RegistosComTempoValido
    , TotalRegistos - RegistosComTempoValido AS [Registos sem tempo valido]

    , CAST(ROUND(MediaColheitaRegisto, 0) AS INT) AS [Media Colheita ao Registo]
    , CAST(ROUND(MediaRegistoProcessamento, 0) AS INT) AS [Media Registo ao Processamento]
    , CAST(ROUND(MediaProcessamentoValidacao, 0) AS INT) AS [Media Processamento a Validacao]
    , CAST(ROUND(MediaTotal, 0) AS INT) AS [Media Total Colheita a Validacao]

    , CASE
          WHEN Ordem = 1 THEN CAST(@MGITTargetDias AS VARCHAR(3)) + ' dias (do registo ao resultado final)'
          WHEN Ordem = 2 THEN CAST(@LJTargetDias AS VARCHAR(3)) + ' dias (do registo ao resultado final)'
          WHEN Ordem = 3 THEN CAST(@LPATargetDias AS VARCHAR(3)) + ' dias (da colheita ao resultado)'
          WHEN Ordem = 4 THEN CAST(@LPATargetDias AS VARCHAR(3)) + ' dias após cultura positiva'
          WHEN Ordem = 5 THEN CAST(@TSATargetDias AS VARCHAR(3)) + ' dias após cultura positiva'
      END AS Target

    , CAST(ROUND(MediaTarget, 0) AS INT) AS [Media do tempo avaliado para o Target]
    , MaiorTempoTarget AS [Maior tempo para o Target]
    , RegistosForaPrazo AS [Registos fora do prazo]

    , CAST
      (
          100.0 * RegistosForaPrazo
          / NULLIF(RegistosComTempoValido, 0)
          AS DECIMAL(6,2)
      ) AS [Percentagem fora do prazo]

    , CASE
          WHEN MediaTarget IS NULL THEN 'Sem dados'
          WHEN MediaTarget <= TargetDias THEN 'Cumprido'
          ELSE 'Não cumprido'
      END AS [Cumprimento do Target pela Media]

    , CASE
          WHEN RegistosComTempoValido = 0
              THEN 'Sem tempos validos para avaliar'
          WHEN RegistosForaPrazo > 0
              THEN 'Existem registos fora do prazo'
          WHEN RegistosComTempoValido < TotalRegistos
              THEN 'Sem atrasos nos tempos validos; existem dados em falta'
          ELSE 'Todos os registos dentro do prazo'
      END AS [Alerta de Atrasos]

FROM Estatisticas
ORDER BY Ordem;
