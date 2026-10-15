SET NOCOUNT ON;
/*
 Motivos de rejeição registados na semana, por laboratório e categoria.
 Parâmetros Python: (@WeekStart, @LabName).
 Cada linha de dbo.TBCultura é contada no máximo uma vez.
 O período é o da data de registo: não há data de rejeição confirmada.
 COM/Comentário fica SEPARADO para validação epidemiológica.
 SQL Server 2014.
 */
DECLARE @WeekStart DATE = ?;
DECLARE @LabName NVARCHAR(20) = ?;
DECLARE @WeekEnd DATE = DATEADD(DAY, 7, @WeekStart);
;
WITH Amostras AS (
    SELECT T.ID,
        NULLIF(
            LTRIM(
                RTRIM(CONVERT(NVARCHAR(100), T.LIMSRejectionCode))
            ),
            N''
        ) AS CodigoBruto,
        NULLIF(
            LTRIM(
                RTRIM(CONVERT(NVARCHAR(400), T.LIMSRejectionDesc))
            ),
            N''
        ) AS DescricaoBruta
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
        AND T.RegisteredDateTime >= @WeekStart
        AND T.RegisteredDateTime < @WeekEnd
        AND (
            NULLIF(
                LTRIM(
                    RTRIM(CONVERT(NVARCHAR(100), T.LIMSRejectionCode))
                ),
                N''
            ) IS NOT NULL
            OR NULLIF(
                LTRIM(
                    RTRIM(CONVERT(NVARCHAR(400), T.LIMSRejectionDesc))
                ),
                N''
            ) IS NOT NULL
        )
),
Normalizado AS (
    SELECT A.ID,
        A.CodigoBruto,
        A.DescricaoBruta,
        UPPER(ISNULL(A.CodigoBruto, N'')) COLLATE Latin1_General_CI_AI AS Codigo,
        UPPER(ISNULL(A.DescricaoBruta, N'')) COLLATE Latin1_General_CI_AI AS Descricao
    FROM Amostras A
),
Classificado AS (
    SELECT N.ID,
        N.CodigoBruto,
        N.DescricaoBruta,
        CASE
            /* COM = Comentário. Não assumir rejeição clínica efectiva. */
            WHEN N.Codigo = N'COM'
            OR N.Descricao IN ('COMENTARIO', 'COMMENT') THEN N'Comentário — requer validação'
            WHEN N.Codigo = 'INS'
            OR N.Descricao LIKE '%AMOSTRA INSUFICIENT%'
            OR N.Descricao LIKE '%INSUFFICIENT%SAMPLE%'
            OR N.Descricao LIKE '%INSUFFICIENT%SPECIMEN%' THEN N'Volume insuficiente'
            WHEN N.Codigo IN (N'EMPT', 'NSRP', N'SDBSA')
            OR N.Descricao LIKE '%RECIPIENTE%VAZIO%'
            OR N.Descricao LIKE '%EMPTY%CONTAINER%'
            OR N.Descricao LIKE '%NO%SPECIMEN%RECEIVED%'
            OR N.Descricao LIKE N'%REQUIS%SEM%AMOSTRA%'
            OR N.Descricao LIKE N'%SEM%AMOSTRA%' THEN N'Amostra não recebida / recipiente vazio'
            WHEN N.Codigo IN (N'AMSID', 'AMONR', 'INLS', 'NAME')
            OR N.Descricao LIKE '%SEM%IDENTIFIC%'
            OR N.Descricao LIKE '%NAO%ROTULAD%'
            OR N.Descricao LIKE '%NOT%LABEL%'
            OR N.Descricao LIKE '%INAD%LABEL%'
            OR N.Descricao LIKE '%NAME%NOT%CORRESPOND%' THEN N'Identificação ou rotulagem inadequada'
            WHEN N.Codigo IN (N'AMMC', N'RECDE', 'UNSIC', N'SOIL')
            OR N.Descricao LIKE '%MAL%ACONDICIONAD%'
            OR N.Descricao LIKE '%RECIPIENTE%CHEIO%'
            OR N.Descricao LIKE '%UNSTERILE%CONT%'
            OR N.Descricao LIKE '%CONTENTOR%SUJO%' THEN N'Acondicionamento / recipiente inadequado'
            WHEN N.Codigo IN ('AMCON', 'SAFUN')
            OR N.Descricao LIKE '%CONTAMINAD%'
            OR N.Descricao LIKE '%CONTAMINAT%' THEN N'Amostra contaminada'
            WHEN N.Codigo IN (N'48HRS', N'OLD30')
            OR N.Descricao LIKE N'%MAIS%48%HORA%'
            OR N.Descricao LIKE N'%OLDER%48%H%'
            OR N.Descricao LIKE '%SPECIMEN%OLDER%' THEN N'Amostra fora do prazo'
            WHEN N.Codigo = N'DRY'
            OR N.Descricao LIKE N'%AMOSTRA%SECA%'
            OR N.Descricao LIKE '%DRY%SPECIMEN%' THEN N'Amostra seca'
            WHEN N.Codigo = N'DUPRG'
            OR N.Descricao LIKE N'%DOUBLE%REGISTR%'
            OR N.Descricao LIKE N'%DUPLIC%' THEN N'Registo duplicado'
            WHEN N.Codigo IN (N'FDR', N'OS')
            OR N.Descricao LIKE '%FALTA%REAGENTE%'
            OR N.Descricao LIKE N'%OUT%OF%STOCK%' THEN N'Indisponibilidade de reagentes'
            WHEN N.Codigo = N'ERRTC'
            OR N.Descricao LIKE '%ERRO%TECNIC%'
            OR N.Descricao LIKE '%TECHNICAL%ERROR%' THEN N'Erro técnico'
            WHEN N.Codigo = 'UNS'
            OR N.Descricao LIKE '%UNSUITABLE%TEST%'
            OR N.Descricao LIKE '%INADEQUAD%TEST%' THEN N'Amostra inadequada para o teste'
            WHEN N.Codigo = N'ACC'
            OR N.Descricao LIKE '%ACIDENTE%LABORATORIAL%' THEN N'Acidente laboratorial'
            WHEN N.Codigo = 'SPUN2'
            OR N.Descricao LIKE N'%REPITA%' THEN N'Repetição solicitada'
            WHEN N.Codigo = N'AREGM'
            OR N.Descricao LIKE '%REGISTERED%MANUALLY%' THEN N'Registo manual'
            WHEN N.Codigo = N'REJMF' THEN N'Rejeição MF — por validar'
            WHEN N.Codigo = N'OUTR' THEN N'Outros motivos'
            ELSE N'Outros / não classificados'
        END AS MotivoAgrupado
    FROM Normalizado N
)
SELECT @LabName AS Laboratorio,
    MotivoAgrupado,
    COALESCE(CodigoBruto, N'(sem código)') AS CodigoOriginal,
    COALESCE(DescricaoBruta, N'(sem descrição)') AS DescricaoOriginal,
    COUNT_BIG(*) AS Quantidade
FROM Classificado
GROUP BY MotivoAgrupado,
    CodigoBruto,
    DescricaoBruta
ORDER BY COUNT_BIG(*) DESC,
    MotivoAgrupado,
    CodigoBruto,
    DescricaoBruta;