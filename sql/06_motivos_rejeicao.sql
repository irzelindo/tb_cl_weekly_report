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

;WITH Amostras AS
(
    SELECT
        T.ID,
        NULLIF(LTRIM(RTRIM(CONVERT(NVARCHAR(100), T.LIMSRejectionCode))), N'') AS CodigoBruto,
        NULLIF(LTRIM(RTRIM(CONVERT(NVARCHAR(400), T.LIMSRejectionDesc))), N'') AS DescricaoBruta
    FROM dbo.TBCultura AS T
    CROSS APPLY
    (
        SELECT UPPER(CONVERT(NVARCHAR(255), COALESCE(
            NULLIF(LTRIM(RTRIM(T.TestingFacilityName)), ''),
            NULLIF(LTRIM(RTRIM(T.LIMSFacilityName)), ''), N''
        ))) COLLATE Latin1_General_CI_AI AS NomeBruto
    ) AS LabRaw
    CROSS APPLY
    (
        SELECT CASE
            WHEN LabRaw.NomeBruto LIKE N'%CARMELO%' THEN N'Carmelo'
            WHEN LabRaw.NomeBruto LIKE N'%NAMPULA%' THEN N'Nampula'
            WHEN LabRaw.NomeBruto LIKE N'%PTB%INS%'
              OR LabRaw.NomeBruto IN (N'TB LABORATORIO', N'TB LABORATÓRIO')
                THEN N'Maputo'
            WHEN LabRaw.NomeBruto LIKE N'%BEIRA%' THEN N'Beira'
            ELSE NULL
        END AS Nome
    ) AS Lab
    WHERE Lab.Nome = @LabName
      AND T.RegisteredDateTime >= @WeekStart
      AND T.RegisteredDateTime < @WeekEnd
      AND (
           NULLIF(LTRIM(RTRIM(CONVERT(NVARCHAR(100), T.LIMSRejectionCode))), N'') IS NOT NULL
        OR NULLIF(LTRIM(RTRIM(CONVERT(NVARCHAR(400), T.LIMSRejectionDesc))), N'') IS NOT NULL
      )
),
Normalizado AS
(
    SELECT
        A.ID,
        A.CodigoBruto,
        A.DescricaoBruta,
        UPPER(ISNULL(A.CodigoBruto, N'')) COLLATE Latin1_General_CI_AI AS Codigo,
        UPPER(ISNULL(A.DescricaoBruta, N'')) COLLATE Latin1_General_CI_AI AS Descricao
    FROM Amostras A
),
Classificado AS
(
    SELECT
        N.ID,
        N.CodigoBruto,
        N.DescricaoBruta,
        CASE
            /* COM = Comentário. Não assumir rejeição clínica efectiva. */
            WHEN N.Codigo = N'COM'
              OR N.Descricao IN (N'COMENTARIO', N'COMMENT')
                THEN N'Comentário — requer validação'

            WHEN N.Codigo = N'INS'
              OR N.Descricao LIKE N'%AMOSTRA INSUFICIENT%'
              OR N.Descricao LIKE N'%INSUFFICIENT%SAMPLE%'
              OR N.Descricao LIKE N'%INSUFFICIENT%SPECIMEN%'
                THEN N'Volume insuficiente'

            WHEN N.Codigo IN (N'EMPT', N'NSRP', N'SDBSA')
              OR N.Descricao LIKE N'%RECIPIENTE%VAZIO%'
              OR N.Descricao LIKE N'%EMPTY%CONTAINER%'
              OR N.Descricao LIKE N'%NO%SPECIMEN%RECEIVED%'
              OR N.Descricao LIKE N'%REQUIS%SEM%AMOSTRA%'
              OR N.Descricao LIKE N'%SEM%AMOSTRA%'
                THEN N'Amostra não recebida / recipiente vazio'

            WHEN N.Codigo IN (N'AMSID', N'AMONR', N'INLS', N'NAME')
              OR N.Descricao LIKE N'%SEM%IDENTIFIC%'
              OR N.Descricao LIKE N'%NAO%ROTULAD%'
              OR N.Descricao LIKE N'%NOT%LABEL%'
              OR N.Descricao LIKE N'%INAD%LABEL%'
              OR N.Descricao LIKE N'%NAME%NOT%CORRESPOND%'
                THEN N'Identificação ou rotulagem inadequada'

            WHEN N.Codigo IN (N'AMMC', N'RECDE', N'UNSIC', N'SOIL')
              OR N.Descricao LIKE N'%MAL%ACONDICIONAD%'
              OR N.Descricao LIKE N'%RECIPIENTE%CHEIO%'
              OR N.Descricao LIKE N'%UNSTERILE%CONT%'
              OR N.Descricao LIKE N'%CONTENTOR%SUJO%'
                THEN N'Acondicionamento / recipiente inadequado'

            WHEN N.Codigo IN (N'AMCON', N'SAFUN')
              OR N.Descricao LIKE N'%CONTAMINAD%'
              OR N.Descricao LIKE N'%CONTAMINAT%'
                THEN N'Amostra contaminada'

            WHEN N.Codigo IN (N'48HRS', N'OLD30')
              OR N.Descricao LIKE N'%MAIS%48%HORA%'
              OR N.Descricao LIKE N'%OLDER%48%H%'
              OR N.Descricao LIKE N'%SPECIMEN%OLDER%'
                THEN N'Amostra fora do prazo'

            WHEN N.Codigo = N'DRY'
              OR N.Descricao LIKE N'%AMOSTRA%SECA%'
              OR N.Descricao LIKE N'%DRY%SPECIMEN%'
                THEN N'Amostra seca'

            WHEN N.Codigo = N'DUPRG'
              OR N.Descricao LIKE N'%DOUBLE%REGISTR%'
              OR N.Descricao LIKE N'%DUPLIC%'
                THEN N'Registo duplicado'

            WHEN N.Codigo IN (N'FDR', N'OS')
              OR N.Descricao LIKE N'%FALTA%REAGENTE%'
              OR N.Descricao LIKE N'%OUT%OF%STOCK%'
                THEN N'Indisponibilidade de reagentes'

            WHEN N.Codigo = N'ERRTC'
              OR N.Descricao LIKE N'%ERRO%TECNIC%'
              OR N.Descricao LIKE N'%TECHNICAL%ERROR%'
                THEN N'Erro técnico'

            WHEN N.Codigo = N'UNS'
              OR N.Descricao LIKE N'%UNSUITABLE%TEST%'
              OR N.Descricao LIKE N'%INADEQUAD%TEST%'
                THEN N'Amostra inadequada para o teste'

            WHEN N.Codigo = N'ACC'
              OR N.Descricao LIKE N'%ACIDENTE%LABORATORIAL%'
                THEN N'Acidente laboratorial'

            WHEN N.Codigo = N'SPUN2'
              OR N.Descricao LIKE N'%REPITA%'
                THEN N'Repetição solicitada'

            WHEN N.Codigo = N'AREGM'
              OR N.Descricao LIKE N'%REGISTERED%MANUALLY%'
                THEN N'Registo manual'

            WHEN N.Codigo = N'REJMF'
                THEN N'Rejeição MF — por validar'

            WHEN N.Codigo = N'OUTR'
                THEN N'Outros motivos'

            ELSE N'Outros / não classificados'
        END AS MotivoAgrupado
    FROM Normalizado N
)
SELECT
    @LabName AS Laboratorio,
    MotivoAgrupado,
    COALESCE(CodigoBruto, N'(sem código)') AS CodigoOriginal,
    COALESCE(DescricaoBruta, N'(sem descrição)') AS DescricaoOriginal,
    COUNT_BIG(*) AS Quantidade
FROM Classificado
GROUP BY MotivoAgrupado, CodigoBruto, DescricaoBruta
ORDER BY COUNT_BIG(*) DESC, MotivoAgrupado, CodigoBruto, DescricaoBruta;
