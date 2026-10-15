SET NOCOUNT ON;
/* Rejeitadas: pelo menos um campo de rejeição preenchido. Não inferir método MGIT/LJ onde não há resultado/marca de método. */

/*
Fluxo semanal + backlog à data de corte. SQL Server 2014.
Parâmetros: (1) semana inicial, (2) laboratório canónico.
Rejeições: NULL até confirmar um campo/código inequívoco de rejeição na TBCultura.
As datas dos resultados permitem reconstruir o backlog histórico quando preenchidas.
*/
DECLARE @WeekStart DATE = ?;
DECLARE @LabName NVARCHAR(20) = ?;
DECLARE @WeekEnd DATE = DATEADD(DAY, 7, @WeekStart);
DECLARE @LastDay DATE = DATEADD(DAY, -1, @WeekEnd);

;WITH Origem AS (
    SELECT
        T.ID, T.HL7SexCode, T.RegisteredDateTime, T.AnalysisDateTime,
        T.AuthorisedDateTime, T.HL7ResultStatusCode,
        T.LIMSRejectionCode, T.LIMSRejectionDesc,
        T.LiquidCultureResult, T.SolidCultureResult,
        T.FinalMGITCultureResult, T.FinalLJCultureResult,
        T.FinalMGITResultDate, T.FinalLJResultDate,
        T.LPAResultDate, T.LPA2ResultDate, T.LPACMResultDate,
        T.TSAResultDate, T.TSA2ResultDate, T.TBIDResultDate,
        Resultados.PrimeiroResultado
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
    OUTER APPLY (
        SELECT MIN(V.DataResultado) AS PrimeiroResultado
        FROM (VALUES
            (TRY_CONVERT(DATE, NULLIF(LTRIM(RTRIM(T.FinalMGITResultDate)), ''), 103)),
            (TRY_CONVERT(DATE, NULLIF(LTRIM(RTRIM(T.FinalLJResultDate)), ''), 103)),
            (TRY_CONVERT(DATE, NULLIF(LTRIM(RTRIM(T.TBIDResultDate)), ''), 103)),
            (TRY_CONVERT(DATE, NULLIF(LTRIM(RTRIM(T.LPAResultDate)), ''), 103)),
            (TRY_CONVERT(DATE, NULLIF(LTRIM(RTRIM(T.LPA2ResultDate)), ''), 103)),
            (TRY_CONVERT(DATE, NULLIF(LTRIM(RTRIM(T.LPACMResultDate)), ''), 103)),
            (TRY_CONVERT(DATE, NULLIF(LTRIM(RTRIM(T.TSAResultDate)), ''), 103)),
            (TRY_CONVERT(DATE, NULLIF(LTRIM(RTRIM(T.TSA2ResultDate)), ''), 103))
        ) AS V(DataResultado)
    ) AS Resultados
    WHERE Lab.Nome = @LabName
      AND T.RegisteredDateTime IS NOT NULL
      AND T.RegisteredDateTime < @WeekEnd
),
Flags AS (
    SELECT O.*,
        CASE WHEN O.RegisteredDateTime >= @WeekStart
              AND O.RegisteredDateTime < @WeekEnd
              AND (
                NULLIF(LTRIM(RTRIM(O.LIMSRejectionCode)), '') IS NOT NULL
                OR NULLIF(LTRIM(RTRIM(O.LIMSRejectionDesc)), '') IS NOT NULL
              )
          THEN 1 ELSE 0 END AS RejeitadaNaSemana,
        CASE WHEN O.RegisteredDateTime >= @WeekStart
               AND O.RegisteredDateTime < @WeekEnd THEN 1 ELSE 0 END AS RegistadaNaSemana,
        CASE WHEN O.AnalysisDateTime >= @WeekStart
               AND O.AnalysisDateTime < @WeekEnd THEN 1 ELSE 0 END AS TestadaNaSemana,
        CASE WHEN O.HL7ResultStatusCode = 'F'
               AND O.AuthorisedDateTime >= @WeekStart
               AND O.AuthorisedDateTime < @WeekEnd THEN 1 ELSE 0 END AS ValidadaNaSemana,
        CASE WHEN
          (O.PrimeiroResultado IS NOT NULL AND O.PrimeiroResultado < @WeekEnd)
          OR (O.HL7ResultStatusCode = 'F' AND O.AuthorisedDateTime < @WeekEnd)
          THEN 0 ELSE 1 END AS PendenteNoCorte,
        CASE WHEN (O.LiquidCultureResult IS NOT NULL AND LTRIM(RTRIM(O.LiquidCultureResult)) <> '')
              OR (O.FinalMGITCultureResult IS NOT NULL AND LTRIM(RTRIM(O.FinalMGITCultureResult)) <> '')
              THEN 1 ELSE 0 END AS TemMGIT,
        CASE WHEN (O.SolidCultureResult IS NOT NULL AND LTRIM(RTRIM(O.SolidCultureResult)) <> '')
              OR (O.FinalLJCultureResult IS NOT NULL AND LTRIM(RTRIM(O.FinalLJCultureResult)) <> '')
              THEN 1 ELSE 0 END AS TemLJ,
        CASE WHEN (O.LiquidCultureResult LIKE 'Contamin%'
                OR O.FinalMGITCultureResult LIKE 'Contamin%') THEN 1 ELSE 0 END AS ContaminadaMGIT,
        CASE WHEN (O.SolidCultureResult LIKE 'Contamin%'
                OR O.FinalLJCultureResult LIKE 'Contamin%') THEN 1 ELSE 0 END AS ContaminadaLJ,
        DATEDIFF(DAY, CAST(O.RegisteredDateTime AS DATE), @LastDay) AS IdadeDias
    FROM Origem O
)
SELECT
    @LabName AS Laboratorio,
    X.Codigo,
    SUM(CASE WHEN X.Flag = 1 AND F.HL7SexCode = 'M' THEN 1 ELSE 0 END) AS Homens,
    SUM(CASE WHEN X.Flag = 1 AND F.HL7SexCode = 'F' THEN 1 ELSE 0 END) AS Mulheres,
    SUM(CASE WHEN X.Flag = 1
              AND (F.HL7SexCode NOT IN ('M','F') OR F.HL7SexCode IS NULL)
              THEN 1 ELSE 0 END) AS NaoEspecificado,
    SUM(X.Flag) AS Total,
    CASE WHEN X.Codigo = 'Pendentes' THEN SUM(CASE WHEN X.Flag = 1 AND F.IdadeDias BETWEEN 0 AND 7 THEN 1 ELSE 0 END) END AS [0_7],
    CASE WHEN X.Codigo = 'Pendentes' THEN SUM(CASE WHEN X.Flag = 1 AND F.IdadeDias BETWEEN 8 AND 44 THEN 1 ELSE 0 END) END AS [8_44],
    CASE WHEN X.Codigo = 'Pendentes' THEN SUM(CASE WHEN X.Flag = 1 AND F.IdadeDias BETWEEN 45 AND 60 THEN 1 ELSE 0 END) END AS [45_60],
    CASE WHEN X.Codigo = 'Pendentes' THEN SUM(CASE WHEN X.Flag = 1 AND F.IdadeDias > 60 THEN 1 ELSE 0 END) END AS [Mais60],
    CASE WHEN X.Codigo = 'Pendentes' THEN MAX(CASE WHEN X.Flag = 1 THEN F.IdadeDias END) END AS MaisAntigaDias
FROM Flags AS F
CROSS APPLY (VALUES
 ('Recebidas', F.RegistadaNaSemana),
 ('Testadas', F.TestadaNaSemana),
 ('Validadas', F.ValidadaNaSemana),
 ('Rejeitadas', F.RejeitadaNaSemana),
 ('Pendentes', F.PendenteNoCorte),
 ('Contaminadas', CASE WHEN F.AnalysisDateTime >= @WeekStart AND F.AnalysisDateTime < @WeekEnd
                     AND (F.ContaminadaMGIT = 1 OR F.ContaminadaLJ = 1) THEN 1 ELSE 0 END),
 ('MGIT_Registadas', CASE WHEN F.TemMGIT = 1 THEN F.RegistadaNaSemana ELSE 0 END),
 ('MGIT_Testadas', CASE WHEN F.TemMGIT = 1 THEN F.TestadaNaSemana ELSE 0 END),
 ('MGIT_Validadas', CASE WHEN F.TemMGIT = 1 THEN F.ValidadaNaSemana ELSE 0 END),
 ('MGIT_Rejeitadas', CASE WHEN F.TemMGIT = 1 THEN F.RejeitadaNaSemana ELSE 0 END),
 ('MGIT_Pendentes', CASE WHEN F.TemMGIT = 1 THEN F.PendenteNoCorte ELSE 0 END),
 ('MGIT_Contaminadas', CASE WHEN F.TemMGIT = 1 AND F.AnalysisDateTime >= @WeekStart
             AND F.AnalysisDateTime < @WeekEnd THEN F.ContaminadaMGIT ELSE 0 END),
 ('LJ_Registadas', CASE WHEN F.TemLJ = 1 THEN F.RegistadaNaSemana ELSE 0 END),
 ('LJ_Testadas', CASE WHEN F.TemLJ = 1 THEN F.TestadaNaSemana ELSE 0 END),
 ('LJ_Validadas', CASE WHEN F.TemLJ = 1 THEN F.ValidadaNaSemana ELSE 0 END),
 ('LJ_Rejeitadas', CASE WHEN F.TemLJ = 1 THEN F.RejeitadaNaSemana ELSE 0 END),
 ('LJ_Pendentes', CASE WHEN F.TemLJ = 1 THEN F.PendenteNoCorte ELSE 0 END),
 ('LJ_Contaminadas', CASE WHEN F.TemLJ = 1 AND F.AnalysisDateTime >= @WeekStart
             AND F.AnalysisDateTime < @WeekEnd THEN F.ContaminadaLJ ELSE 0 END)
) AS X(Codigo, Flag)
GROUP BY X.Codigo
ORDER BY X.Codigo;
