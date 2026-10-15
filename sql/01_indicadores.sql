SET NOCOUNT ON;
/* ============================================================
 PARÂMETROS RECEBIDOS DO PYTHON
 ============================================================ */
DECLARE @WeekStart DATE = ?;
DECLARE @LabName NVARCHAR(20) = ?;
DECLARE @WeekEnd DATE = DATEADD(DAY, 7, @WeekStart);
;
WITH Base AS (
    SELECT T.ID,
        T.RequestID,
        T.HL7SexCode,
        T.AnalysisDateTime,
        T.LiquidCultureResult,
        T.SolidCultureResult,
        T.FinalMGITCultureResult,
        T.FinalLJCultureResult,
        T.TBIDResult,
        T.LPAComplexMTBResult,
        T.LPAIsoniazidResult,
        T.LPARifampicinResult,
        T.LPA2ComplexMTBResult,
        T.LPA2FluoroquinoloneResult,
        T.LPA2KACResult,
        T.LPA2KACVResult,
        T.LPA2KCVResult,
        T.LPA2LowLevelKanamycinResult,
        T.LPACMResult,
        T.TSAEthambutolResult,
        T.TSAIsoniazidResult,
        T.TSAPyrazinamideResult,
        T.TSARifampicinResult,
        T.TSAStreptomycinResult,
        T.TSA2AmikacinResult,
        T.TSA2BedaquilineResult,
        T.TSA2ClofazimineResult,
        T.TSA2CapreomycinResult,
        T.TSA2EthionamideResult,
        T.TSA2KanamycinResult,
        T.TSA2LevofloxacinResult,
        T.TSA2LinezolidResult,
        T.TSA2MoxifloxacinResult,
        T.TSA2OfloxacinResult,
        T.TSA2ProthionamideResult
        /* =====================================================
         CULTURA LÍQUIDA (MGIT)
         ===================================================== */
,
        CASE
            WHEN NULLIF(LTRIM(RTRIM(T.LiquidCultureResult)), '') IS NOT NULL
            OR NULLIF(LTRIM(RTRIM(T.FinalMGITCultureResult)), '') IS NOT NULL THEN 1
            ELSE 0
        END AS HasMGIT,
        CASE
            WHEN LTRIM(RTRIM(T.LiquidCultureResult)) IN (
                'Contaminated*',
                'Contaminada*',
                'Contaminated',
                'Contaminado'
            )
            OR LTRIM(RTRIM(T.FinalMGITCultureResult)) IN (
                'Contaminated',
                'Contaminado',
                'Contaminated - Possible NTM'
            ) THEN 1
            ELSE 0
        END AS MGITContaminado,
        CASE
            WHEN LTRIM(RTRIM(T.FinalMGITCultureResult)) IN (
                'Positive for MTB Complex',
                'Positivo para Complexo MTB',
                'Positivo para CMTB',
                'C. Mycobacterium tuberculosis*'
            ) THEN 1
            ELSE 0
        END AS MGIT_CMTB,
        CASE
            WHEN LTRIM(RTRIM(T.FinalMGITCultureResult)) IN (
                'Non-tuberculosis Mycobacteria',
                'Positive for NTM',
                'Mycobacteria NÃO Tuberculosis',
                'Micobacterias Nao TB',
                'MNT'
            ) THEN 1
            ELSE 0
        END AS MGIT_MNT,
        CASE
            WHEN LTRIM(RTRIM(T.FinalMGITCultureResult)) IN (
                'Negative for MTB Complex',
                'Negativo para Complexo MTB',
                'Negativo para CMTB',
                'Negative',
                'Negativo'
            ) THEN 1
            ELSE 0
        END AS MGIT_Negativo
        /* =====================================================
         CULTURA SÓLIDA (LJ)
         ===================================================== */
,
        CASE
            WHEN NULLIF(LTRIM(RTRIM(T.SolidCultureResult)), '') IS NOT NULL
            OR NULLIF(LTRIM(RTRIM(T.FinalLJCultureResult)), '') IS NOT NULL THEN 1
            ELSE 0
        END AS HasLJ,
        CASE
            WHEN LTRIM(RTRIM(T.SolidCultureResult)) IN (
                'Contaminated*',
                'Contaminada*',
                'Contaminated',
                'Contaminado'
            )
            OR LTRIM(RTRIM(T.FinalLJCultureResult)) IN ('Contaminated', 'Contaminado') THEN 1
            ELSE 0
        END AS LJContaminado,
        CASE
            WHEN LTRIM(RTRIM(T.FinalLJCultureResult)) IN (
                'Positive for MTB Complex',
                'Positivo para Complexo MTB',
                'Positivo para CMTB',
                'C. Mycobacterium tuberculosis*'
            ) THEN 1
            ELSE 0
        END AS LJ_CMTB,
        CASE
            WHEN LTRIM(RTRIM(T.FinalLJCultureResult)) IN (
                'Non-tuberculosis Mycobacteria',
                'Positive for NTM',
                'Mycobacteria NÃO Tuberculosis',
                'Micobacterias Nao TB',
                'MNT'
            ) THEN 1
            ELSE 0
        END AS LJ_MNT,
        CASE
            WHEN LTRIM(RTRIM(T.FinalLJCultureResult)) IN (
                'Negative for MTB Complex',
                'Negativo para Complexo MTB',
                'Negativo para CMTB',
                'Negative',
                'Negativo'
            ) THEN 1
            ELSE 0
        END AS LJ_Negativo
        /* =====================================================
         LPA 1 / IDENTIFICAÇÃO
         ===================================================== */
,
        CASE
            WHEN NULLIF(LTRIM(RTRIM(T.LPAComplexMTBResult)), '') IS NOT NULL
            OR NULLIF(LTRIM(RTRIM(T.LPAIsoniazidResult)), '') IS NOT NULL
            OR NULLIF(LTRIM(RTRIM(T.LPARifampicinResult)), '') IS NOT NULL THEN 1
            ELSE 0
        END AS HasLPA1,
        CASE
            WHEN LTRIM(RTRIM(T.LPAComplexMTBResult)) IN ('Detectado', 'Detected') THEN 1
            ELSE 0
        END AS LPA1_CMTB,
        CASE
            /* MNT só pode ser atribuída à secção LPA1 quando
             o próprio registo possui LPA1 realizado.
             
             TBIDResult é usado apenas como evidência explícita
             de MNT associada à mesma amostra; NÃO interpretamos
             'Não detectado' no LPA1 como MNT. */
            WHEN (
                NULLIF(LTRIM(RTRIM(T.LPAComplexMTBResult)), '') IS NOT NULL
                OR NULLIF(LTRIM(RTRIM(T.LPAIsoniazidResult)), '') IS NOT NULL
                OR NULLIF(LTRIM(RTRIM(T.LPARifampicinResult)), '') IS NOT NULL
            )
            AND LTRIM(RTRIM(T.TBIDResult)) IN (
                'Micobacterias Nao TB',
                'Mycobacteria NÃO Tuberculosis',
                '^BNon TB Complex^B',
                '^BMicobacterias Nao TB^B',
                'MNT'
            ) THEN 1
            ELSE 0
        END AS LPA1_MNT
        /* =====================================================
         LPA 2 / LPA CM
         ===================================================== */
,
        CASE
            WHEN NULLIF(LTRIM(RTRIM(T.LPA2ComplexMTBResult)), '') IS NOT NULL
            OR NULLIF(LTRIM(RTRIM(T.LPA2FluoroquinoloneResult)), '') IS NOT NULL
            OR NULLIF(LTRIM(RTRIM(T.LPA2KACResult)), '') IS NOT NULL
            OR NULLIF(LTRIM(RTRIM(T.LPA2KACVResult)), '') IS NOT NULL
            OR NULLIF(LTRIM(RTRIM(T.LPA2KCVResult)), '') IS NOT NULL
            OR NULLIF(LTRIM(RTRIM(T.LPA2LowLevelKanamycinResult)), '') IS NOT NULL THEN 1
            ELSE 0
        END AS HasLPA2,
        CASE
            WHEN NULLIF(LTRIM(RTRIM(T.LPACMResult)), '') IS NOT NULL THEN 1
            ELSE 0
        END AS HasLPACM
        /* =====================================================
         TSA 1 / TSA 2
         ===================================================== */
,
        CASE
            WHEN NULLIF(LTRIM(RTRIM(T.TSAEthambutolResult)), '') IS NOT NULL
            OR NULLIF(LTRIM(RTRIM(T.TSAIsoniazidResult)), '') IS NOT NULL
            OR NULLIF(LTRIM(RTRIM(T.TSAPyrazinamideResult)), '') IS NOT NULL
            OR NULLIF(LTRIM(RTRIM(T.TSARifampicinResult)), '') IS NOT NULL
            OR NULLIF(LTRIM(RTRIM(T.TSAStreptomycinResult)), '') IS NOT NULL THEN 1
            ELSE 0
        END AS HasTSA1,
        CASE
            WHEN NULLIF(LTRIM(RTRIM(T.TSA2AmikacinResult)), '') IS NOT NULL
            OR NULLIF(LTRIM(RTRIM(T.TSA2BedaquilineResult)), '') IS NOT NULL
            OR NULLIF(LTRIM(RTRIM(T.TSA2ClofazimineResult)), '') IS NOT NULL
            OR NULLIF(LTRIM(RTRIM(T.TSA2CapreomycinResult)), '') IS NOT NULL
            OR NULLIF(LTRIM(RTRIM(T.TSA2EthionamideResult)), '') IS NOT NULL
            OR NULLIF(LTRIM(RTRIM(T.TSA2KanamycinResult)), '') IS NOT NULL
            OR NULLIF(LTRIM(RTRIM(T.TSA2LevofloxacinResult)), '') IS NOT NULL
            OR NULLIF(LTRIM(RTRIM(T.TSA2LinezolidResult)), '') IS NOT NULL
            OR NULLIF(LTRIM(RTRIM(T.TSA2MoxifloxacinResult)), '') IS NOT NULL
            OR NULLIF(LTRIM(RTRIM(T.TSA2OfloxacinResult)), '') IS NOT NULL
            OR NULLIF(LTRIM(RTRIM(T.TSA2ProthionamideResult)), '') IS NOT NULL THEN 1
            ELSE 0
        END AS HasTSA2
        /* =====================================================
         CLASSIFICAÇÃO DE RESISTÊNCIA
         ===================================================== */
,
        CASE
            WHEN LTRIM(RTRIM(T.LPARifampicinResult)) IN ('Resistant', 'Resistente')
            OR LTRIM(RTRIM(T.TSARifampicinResult)) IN ('Resistant', 'Resistente') THEN 1
            ELSE 0
        END AS RifampicinaResistente,
        CASE
            WHEN LTRIM(RTRIM(T.LPA2FluoroquinoloneResult)) IN ('Resistant', 'Resistente')
            OR LTRIM(RTRIM(T.TSA2LevofloxacinResult)) IN ('Resistant', 'Resistente')
            OR LTRIM(RTRIM(T.TSA2MoxifloxacinResult)) IN ('Resistant', 'Resistente')
            OR LTRIM(RTRIM(T.TSA2OfloxacinResult)) IN ('Resistant', 'Resistente') THEN 1
            ELSE 0
        END AS FluoroquinolonaResistente,
        CASE
            WHEN LTRIM(RTRIM(T.TSA2BedaquilineResult)) IN ('Resistant', 'Resistente')
            OR LTRIM(RTRIM(T.TSA2LinezolidResult)) IN ('Resistant', 'Resistente') THEN 1
            ELSE 0
        END AS BedaquilinaLinezolidaResistente
    FROM dbo.TBCultura AS T
        CROSS APPLY (
            SELECT UPPER(
                    CONVERT(
                        NVARCHAR(255),
                        COALESCE(
                            NULLIF(LTRIM(RTRIM(T.TestingFacilityName)), ''),
                            NULLIF(LTRIM(RTRIM(T.LIMSFacilityName)), ''),
                            N''
                        )
                    )
                ) COLLATE Latin1_General_CI_AI AS NomeBruto
        ) AS LabRaw
        CROSS APPLY (
            SELECT CASE
                    WHEN LabRaw.NomeBruto LIKE '%CARMELO%'
                    OR LabRaw.NomeBruto IN ('TB LABORATORIO', 'TB LABORATÓRIO') THEN 'Carmelo'
                    WHEN LabRaw.NomeBruto LIKE '%NAMPULA%' THEN 'Nampula'
                    WHEN LabRaw.NomeBruto LIKE '%PTB%INS%' THEN 'Maputo'
                    WHEN LabRaw.NomeBruto LIKE '%BEIRA%' THEN 'Beira'
                    ELSE NULL
                END AS Nome
        ) AS Lab
    WHERE Lab.Nome = @LabName
        AND T.HL7ResultStatusCode = 'F'
        AND T.AnalysisDateTime >= @WeekStart
        AND T.AnalysisDateTime < @WeekEnd
),
Dados AS (
    SELECT B.*,
        CASE
            WHEN B.RifampicinaResistente = 1 THEN 1
            ELSE 0
        END AS RR_MR,
        CASE
            WHEN B.RifampicinaResistente = 1
            AND B.FluoroquinolonaResistente = 1 THEN 1
            ELSE 0
        END AS Pre_XR,
        CASE
            WHEN B.RifampicinaResistente = 1
            AND B.FluoroquinolonaResistente = 1
            AND B.BedaquilinaLinezolidaResistente = 1 THEN 1
            ELSE 0
        END AS XR
    FROM Base AS B
)
/* ============================================================
 1. TOTAL DE AMOSTRAS PROCESSADAS
 ============================================================ */
SELECT @WeekStart AS DataInicio,
    DATEADD(DAY, -1, @WeekEnd) AS DataFim,
    1 AS Ordem,
    '# Total de amostras processadas' AS Indicador,
    SUM(
        CASE
            WHEN HL7SexCode = 'M' THEN 1
            ELSE 0
        END
    ) AS Homens,
    SUM(
        CASE
            WHEN HL7SexCode = 'F' THEN 1
            ELSE 0
        END
    ) AS Mulheres,
    SUM(
        CASE
            WHEN ISNULL(LTRIM(RTRIM(HL7SexCode)), '') NOT IN ('M', 'F') THEN 1
            ELSE 0
        END
    ) AS NaoEspecificado,
    SUM(HasMGIT) AS Liquido_MGIT,
    SUM(HasLJ) AS Solido_LJ,
    SUM(MGIT_CMTB) AS MGIT_CMTB,
    SUM(MGIT_MNT) AS MGIT_MNT,
    SUM(MGIT_Negativo) AS MGIT_Negativo,
    SUM(LJ_CMTB) AS LJ_CMTB,
    SUM(LJ_MNT) AS LJ_MNT,
    SUM(LJ_Negativo) AS LJ_Negativo,
    SUM(LPA1_CMTB) AS LPA1_CMTB,
    SUM(LPA1_MNT) AS LPA1_MNT
    /* Número total de amostras/registos com LPA1 realizado.
     Não é calculado como CMTB + MNT, porque podem existir
     resultados LPA1 válidos que não pertençam a essas duas
     categorias resumidas. */
,
    SUM(HasLPA1) AS LPA1_Total,
    SUM(HasLPA2) AS LPA2,
    SUM(HasLPACM) AS LPA_CM,
    SUM(HasTSA1) AS TSA1,
    SUM(HasTSA2) AS TSA2
    /* Total de amostras/registos com TSA realizado.
     Conta cada registo apenas uma vez, mesmo quando possui
     simultaneamente TSA1 e TSA2. */
,
    SUM(
        CASE
            WHEN HasTSA1 = 1
            OR HasTSA2 = 1 THEN 1
            ELSE 0
        END
    ) AS TSA_Total,
    COUNT_BIG(*) AS Total
FROM Dados
UNION ALL
/* ============================================================
 2. AMOSTRAS CONTAMINADAS
 ============================================================ */
SELECT @WeekStart,
    DATEADD(DAY, -1, @WeekEnd),
    2,
    '# Total de amostras contaminadas',
    SUM(
        CASE
            WHEN HL7SexCode = 'M'
            AND (
                MGITContaminado = 1
                OR LJContaminado = 1
            ) THEN 1
            ELSE 0
        END
    ),
    SUM(
        CASE
            WHEN HL7SexCode = 'F'
            AND (
                MGITContaminado = 1
                OR LJContaminado = 1
            ) THEN 1
            ELSE 0
        END
    ),
    SUM(
        CASE
            WHEN ISNULL(LTRIM(RTRIM(HL7SexCode)), '') NOT IN ('M', 'F')
            AND (
                MGITContaminado = 1
                OR LJContaminado = 1
            ) THEN 1
            ELSE 0
        END
    ),
    SUM(MGITContaminado),
    SUM(LJContaminado)
    /* Resultado final CMTB/MNT/Negativo não é aplicável a contaminação. */
,
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    NULL
    /* LPA/TSA não são indicadores de contaminação.
     Ordem das colunas:
     LPA1_CMTB, LPA1_MNT, LPA1_Total,
     LPA2, LPA_CM, TSA1, TSA2, TSA_Total. */
,
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    SUM(
        CASE
            WHEN MGITContaminado = 1
            OR LJContaminado = 1 THEN 1
            ELSE 0
        END
    )
FROM Dados
UNION ALL
/* ============================================================
 3. RR/MR
 ============================================================ */
SELECT @WeekStart,
    DATEADD(DAY, -1, @WeekEnd),
    3,
    'Resistente à Rifampicina / Multirresistente (RR/MR)',
    SUM(
        CASE
            WHEN RR_MR = 1
            AND HL7SexCode = 'M' THEN 1
            ELSE 0
        END
    ),
    SUM(
        CASE
            WHEN RR_MR = 1
            AND HL7SexCode = 'F' THEN 1
            ELSE 0
        END
    ),
    SUM(
        CASE
            WHEN RR_MR = 1
            AND ISNULL(LTRIM(RTRIM(HL7SexCode)), '') NOT IN ('M', 'F') THEN 1
            ELSE 0
        END
    )
    /* Colunas de cultura não se aplicam à classificação de resistência. */
,
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    SUM(
        CASE
            WHEN RR_MR = 1
            AND LPA1_CMTB = 1 THEN 1
            ELSE 0
        END
    ),
    NULL,
    SUM(
        CASE
            WHEN RR_MR = 1
            AND HasLPA1 = 1 THEN 1
            ELSE 0
        END
    ) AS LPA1_Total,
    SUM(
        CASE
            WHEN RR_MR = 1
            AND HasLPA2 = 1 THEN 1
            ELSE 0
        END
    ),
    NULL,
    SUM(
        CASE
            WHEN RR_MR = 1
            AND HasTSA1 = 1 THEN 1
            ELSE 0
        END
    ),
    SUM(
        CASE
            WHEN RR_MR = 1
            AND HasTSA2 = 1 THEN 1
            ELSE 0
        END
    ),
    SUM(
        CASE
            WHEN RR_MR = 1
            AND (
                HasTSA1 = 1
                OR HasTSA2 = 1
            ) THEN 1
            ELSE 0
        END
    ) AS TSA_Total,
    SUM(RR_MR)
FROM Dados
UNION ALL
/* ============================================================
 4. PRÉ-XR
 ============================================================ */
SELECT @WeekStart,
    DATEADD(DAY, -1, @WeekEnd),
    4,
    'Pré-XR (RR/MR + resistência a fluoroquinolonas)',
    SUM(
        CASE
            WHEN Pre_XR = 1
            AND HL7SexCode = 'M' THEN 1
            ELSE 0
        END
    ),
    SUM(
        CASE
            WHEN Pre_XR = 1
            AND HL7SexCode = 'F' THEN 1
            ELSE 0
        END
    ),
    SUM(
        CASE
            WHEN Pre_XR = 1
            AND ISNULL(LTRIM(RTRIM(HL7SexCode)), '') NOT IN ('M', 'F') THEN 1
            ELSE 0
        END
    ),
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    SUM(
        CASE
            WHEN Pre_XR = 1
            AND LPA1_CMTB = 1 THEN 1
            ELSE 0
        END
    ),
    NULL,
    SUM(
        CASE
            WHEN Pre_XR = 1
            AND HasLPA1 = 1 THEN 1
            ELSE 0
        END
    ) AS LPA1_Total,
    SUM(
        CASE
            WHEN Pre_XR = 1
            AND HasLPA2 = 1 THEN 1
            ELSE 0
        END
    ),
    NULL,
    SUM(
        CASE
            WHEN Pre_XR = 1
            AND HasTSA1 = 1 THEN 1
            ELSE 0
        END
    ),
    SUM(
        CASE
            WHEN Pre_XR = 1
            AND HasTSA2 = 1 THEN 1
            ELSE 0
        END
    ),
    SUM(
        CASE
            WHEN Pre_XR = 1
            AND (
                HasTSA1 = 1
                OR HasTSA2 = 1
            ) THEN 1
            ELSE 0
        END
    ) AS TSA_Total,
    SUM(Pre_XR)
FROM Dados
UNION ALL
/* ============================================================
 5. XR
 ============================================================ */
SELECT @WeekStart,
    DATEADD(DAY, -1, @WeekEnd),
    5,
    'XR (Pré-XR + resistência a Bedaquilina e/ou Linezolida)',
    SUM(
        CASE
            WHEN XR = 1
            AND HL7SexCode = 'M' THEN 1
            ELSE 0
        END
    ),
    SUM(
        CASE
            WHEN XR = 1
            AND HL7SexCode = 'F' THEN 1
            ELSE 0
        END
    ),
    SUM(
        CASE
            WHEN XR = 1
            AND ISNULL(LTRIM(RTRIM(HL7SexCode)), '') NOT IN ('M', 'F') THEN 1
            ELSE 0
        END
    ),
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    NULL,
    SUM(
        CASE
            WHEN XR = 1
            AND LPA1_CMTB = 1 THEN 1
            ELSE 0
        END
    ),
    NULL,
    SUM(
        CASE
            WHEN XR = 1
            AND HasLPA1 = 1 THEN 1
            ELSE 0
        END
    ) AS LPA1_Total,
    SUM(
        CASE
            WHEN XR = 1
            AND HasLPA2 = 1 THEN 1
            ELSE 0
        END
    ),
    NULL,
    SUM(
        CASE
            WHEN XR = 1
            AND HasTSA1 = 1 THEN 1
            ELSE 0
        END
    ),
    SUM(
        CASE
            WHEN XR = 1
            AND HasTSA2 = 1 THEN 1
            ELSE 0
        END
    ),
    SUM(
        CASE
            WHEN XR = 1
            AND (
                HasTSA1 = 1
                OR HasTSA2 = 1
            ) THEN 1
            ELSE 0
        END
    ) AS TSA_Total,
    SUM(XR)
FROM Dados
ORDER BY Ordem;