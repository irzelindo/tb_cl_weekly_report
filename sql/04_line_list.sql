SET NOCOUNT ON;


/* ============================================================
   PARÂMETROS RECEBIDOS DO PYTHON
   ============================================================ */

DECLARE @WeekStart DATE = ?;
DECLARE @WeekEnd   DATE = DATEADD(DAY, 7, @WeekStart);

/* URL base dos laudos */
DECLARE @BaseLaudoURL VARCHAR(200) =
    'https://api.openldr.org.mz/tb/cl/patients/';



;WITH Dados AS
(
    SELECT
        T.*,


        /* =====================================================
           CÓDIGO DO DISA
           ===================================================== */

            CASE
                WHEN T.RequestID IS NOT NULL
                 AND LEN(LTRIM(RTRIM(T.RequestID))) >= 7

                    THEN SUBSTRING
                         (
                             LTRIM(RTRIM(T.RequestID)),
                             7,
                             LEN(LTRIM(RTRIM(T.RequestID)))
                         )

                ELSE NULL
            END AS CodigoDISA,


        /* =====================================================
           RESULTADO FINAL MGIT
           ===================================================== */

        CASE

            WHEN LTRIM(RTRIM(T.FinalMGITCultureResult)) IN
            (
                'Positive for MTB Complex',
                'Positivo para Complexo MTB',
                'Positivo para CMTB',
                'C. Mycobacterium tuberculosis*'
            )
                THEN 'CMTB'


            WHEN LTRIM(RTRIM(T.FinalMGITCultureResult)) IN
            (
                'Negative for MTB Complex',
                'Negativo para Complexo MTB',
                'Negativo para CMTB'
            )
                THEN 'Negativo'


            WHEN LTRIM(RTRIM(T.FinalMGITCultureResult)) IN
            (
                'Contaminated',
                'Contaminado',
                'Contaminated - Possible NTM'
            )
                THEN 'Contaminado'


            WHEN LTRIM(RTRIM(T.FinalMGITCultureResult)) IN
            (
                'Non-tuberculosis Mycobacteria',
                'Positive for NTM',
                'Mycobacteria NÃO Tuberculosis',
                'Micobacterias Nao TB',
                'MNT'
            )
                THEN 'MNT'


            WHEN NULLIF
                 (
                     LTRIM(RTRIM(T.FinalMGITCultureResult)),
                     ''
                 ) IS NULL
                THEN 'Não realizado'


            ELSE LTRIM(RTRIM(T.FinalMGITCultureResult))

        END AS MGIT_Final,


        /* =====================================================
           RESULTADO FINAL LJ
           ===================================================== */

        CASE

            WHEN LTRIM(RTRIM(T.FinalLJCultureResult)) IN
            (
                'Positive for MTB Complex',
                'Positivo para Complexo MTB',
                'Positivo para CMTB',
                'C. Mycobacterium tuberculosis*'
            )
                THEN 'CMTB'


            WHEN LTRIM(RTRIM(T.FinalLJCultureResult)) IN
            (
                'Negative for MTB Complex',
                'Negativo para Complexo MTB',
                'Negativo para CMTB'
            )
                THEN 'Negativo'


            WHEN LTRIM(RTRIM(T.FinalLJCultureResult)) IN
            (
                'Contaminated',
                'Contaminado'
            )
                THEN 'Contaminado'


            WHEN LTRIM(RTRIM(T.FinalLJCultureResult)) IN
            (
                'Non-tuberculosis Mycobacteria',
                'Positive for NTM',
                'Mycobacteria NÃO Tuberculosis',
                'Micobacterias Nao TB',
                'MNT'
            )
                THEN 'MNT'


            WHEN NULLIF
                 (
                     LTRIM(RTRIM(T.FinalLJCultureResult)),
                     ''
                 ) IS NULL
                THEN 'Não realizado'


            ELSE LTRIM(RTRIM(T.FinalLJCultureResult))

        END AS LJ_Final

    FROM dbo.TBCultura AS T

    WHERE T.HL7ResultStatusCode = 'F'
      AND T.AnalysisDateTime >= @WeekStart
      AND T.AnalysisDateTime <  @WeekEnd
),


Normalizado AS
(
    SELECT
        D.*,


        /* =====================================================
           LPA 1 - ISONIAZIDA
           ===================================================== */

        CASE

            WHEN LTRIM(RTRIM(D.LPAIsoniazidResult))
                 IN ('Susceptible', 'Sensível', 'Sensivel')
                THEN 'Sensível'

            WHEN LTRIM(RTRIM(D.LPAIsoniazidResult))
                 IN ('Resistant', 'Resistente')
                THEN 'Resistente'

            WHEN LTRIM(RTRIM(D.LPAIsoniazidResult))
                 IN ('Intermediate', 'Intermédio')
                THEN 'Intermédio'

            WHEN NULLIF
                 (
                     LTRIM(RTRIM(D.LPAIsoniazidResult)),
                     ''
                 ) IS NULL
                THEN 'Não realizado'

            ELSE LTRIM(RTRIM(D.LPAIsoniazidResult))

        END AS LPA1_Isoniazida,


        /* =====================================================
           LPA 1 - RIFAMPICINA
           ===================================================== */

        CASE

            WHEN LTRIM(RTRIM(D.LPARifampicinResult))
                 IN ('Susceptible', 'Sensível', 'Sensivel')
                THEN 'Sensível'

            WHEN LTRIM(RTRIM(D.LPARifampicinResult))
                 IN ('Resistant', 'Resistente')
                THEN 'Resistente'

            WHEN LTRIM(RTRIM(D.LPARifampicinResult))
                 IN ('Intermediate', 'Intermédio')
                THEN 'Intermédio'

            WHEN NULLIF
                 (
                     LTRIM(RTRIM(D.LPARifampicinResult)),
                     ''
                 ) IS NULL
                THEN 'Não realizado'

            ELSE LTRIM(RTRIM(D.LPARifampicinResult))

        END AS LPA1_Rifampicina,


        /* =====================================================
           LPA CM
           ===================================================== */

        CASE

            WHEN NULLIF
                 (
                     LTRIM(RTRIM(D.LPACMResult)),
                     ''
                 ) IS NULL
                THEN 'Não realizado'

            WHEN LTRIM(RTRIM(D.LPACMResult)) =
                 'Mycobacterium tuberculosis'
                THEN 'Complexo M. tuberculosis'

            WHEN LTRIM(RTRIM(D.LPACMResult)) IN
                 (
                     'NDETE',
                     'Não Detectado',
                     'Negativo'
                 )
                THEN 'Não detectado'

            ELSE LTRIM(RTRIM(D.LPACMResult))

        END AS LPACM_Identificacao,


        /* =====================================================
           TSA 1 - STREPTOMICINA
           ===================================================== */

        CASE

            WHEN LTRIM(RTRIM(D.TSAStreptomycinResult))
                 IN ('Susceptible', 'Sensível', 'Sensivel')
                THEN 'Sensível'

            WHEN LTRIM(RTRIM(D.TSAStreptomycinResult))
                 IN ('Resistant', 'Resistente')
                THEN 'Resistente'

            WHEN LTRIM(RTRIM(D.TSAStreptomycinResult))
                 IN ('Intermediate', 'Intermédio')
                THEN 'Intermédio'

            WHEN NULLIF
                 (
                     LTRIM(RTRIM(D.TSAStreptomycinResult)),
                     ''
                 ) IS NULL
                THEN 'Não realizado'

            ELSE LTRIM(RTRIM(D.TSAStreptomycinResult))

        END AS TSA1_Streptomicina,


        /* =====================================================
           TSA 1 - ISONIAZIDA
           ===================================================== */

        CASE

            WHEN LTRIM(RTRIM(D.TSAIsoniazidResult))
                 IN ('Susceptible', 'Sensível', 'Sensivel')
                THEN 'Sensível'

            WHEN LTRIM(RTRIM(D.TSAIsoniazidResult))
                 IN ('Resistant', 'Resistente')
                THEN 'Resistente'

            WHEN LTRIM(RTRIM(D.TSAIsoniazidResult))
                 IN ('Intermediate', 'Intermédio')
                THEN 'Intermédio'

            WHEN NULLIF
                 (
                     LTRIM(RTRIM(D.TSAIsoniazidResult)),
                     ''
                 ) IS NULL
                THEN 'Não realizado'

            ELSE LTRIM(RTRIM(D.TSAIsoniazidResult))

        END AS TSA1_Isoniazida,


        /* =====================================================
           TSA 1 - RIFAMPICINA
           ===================================================== */

        CASE

            WHEN LTRIM(RTRIM(D.TSARifampicinResult))
                 IN ('Susceptible', 'Sensível', 'Sensivel')
                THEN 'Sensível'

            WHEN LTRIM(RTRIM(D.TSARifampicinResult))
                 IN ('Resistant', 'Resistente')
                THEN 'Resistente'

            WHEN LTRIM(RTRIM(D.TSARifampicinResult))
                 IN ('Intermediate', 'Intermédio')
                THEN 'Intermédio'

            WHEN NULLIF
                 (
                     LTRIM(RTRIM(D.TSARifampicinResult)),
                     ''
                 ) IS NULL
                THEN 'Não realizado'

            ELSE LTRIM(RTRIM(D.TSARifampicinResult))

        END AS TSA1_Rifampicina,


        /* =====================================================
           TSA 1 - ETAMBUTOL
           ===================================================== */

        CASE

            WHEN LTRIM(RTRIM(D.TSAEthambutolResult))
                 IN ('Susceptible', 'Sensível', 'Sensivel')
                THEN 'Sensível'

            WHEN LTRIM(RTRIM(D.TSAEthambutolResult))
                 IN ('Resistant', 'Resistente')
                THEN 'Resistente'

            WHEN LTRIM(RTRIM(D.TSAEthambutolResult))
                 IN ('Intermediate', 'Intermédio')
                THEN 'Intermédio'

            WHEN NULLIF
                 (
                     LTRIM(RTRIM(D.TSAEthambutolResult)),
                     ''
                 ) IS NULL
                THEN 'Não realizado'

            ELSE LTRIM(RTRIM(D.TSAEthambutolResult))

        END AS TSA1_Etambutol,


        /* =====================================================
           TSA 2 - BEDAQUILINA
           ===================================================== */

        CASE

            WHEN LTRIM(RTRIM(D.TSA2BedaquilineResult))
                 IN ('Susceptible', 'Sensível', 'Sensivel')
                THEN 'Sensível'

            WHEN LTRIM(RTRIM(D.TSA2BedaquilineResult))
                 IN ('Resistant', 'Resistente')
                THEN 'Resistente'

            WHEN NULLIF
                 (
                     LTRIM(RTRIM(D.TSA2BedaquilineResult)),
                     ''
                 ) IS NULL
                THEN 'Não realizado'

            ELSE LTRIM(RTRIM(D.TSA2BedaquilineResult))

        END AS TSA2_Bedaquilina,


        /* =====================================================
           TSA 2 - LINEZOLIDA
           ===================================================== */

        CASE

            WHEN LTRIM(RTRIM(D.TSA2LinezolidResult))
                 IN ('Susceptible', 'Sensível', 'Sensivel')
                THEN 'Sensível'

            WHEN LTRIM(RTRIM(D.TSA2LinezolidResult))
                 IN ('Resistant', 'Resistente')
                THEN 'Resistente'

            WHEN NULLIF
                 (
                     LTRIM(RTRIM(D.TSA2LinezolidResult)),
                     ''
                 ) IS NULL
                THEN 'Não realizado'

            ELSE LTRIM(RTRIM(D.TSA2LinezolidResult))

        END AS TSA2_Linezolida,


        /* =====================================================
           TSA 2 - MOXIFLOXACINA
           ===================================================== */

        CASE

            WHEN LTRIM(RTRIM(D.TSA2MoxifloxacinResult))
                 IN ('Susceptible', 'Sensível', 'Sensivel')
                THEN 'Sensível'

            WHEN LTRIM(RTRIM(D.TSA2MoxifloxacinResult))
                 IN ('Resistant', 'Resistente')
                THEN 'Resistente'

            WHEN NULLIF
                 (
                     LTRIM(RTRIM(D.TSA2MoxifloxacinResult)),
                     ''
                 ) IS NULL
                THEN 'Não realizado'

            ELSE LTRIM(RTRIM(D.TSA2MoxifloxacinResult))

        END AS TSA2_Moxifloxacina,


        /* =====================================================
           TSA 2 - CLOFAZIMINA
           ===================================================== */

        CASE

            WHEN LTRIM(RTRIM(D.TSA2ClofazimineResult))
                 IN ('Susceptible', 'Sensível', 'Sensivel')
                THEN 'Sensível'

            WHEN LTRIM(RTRIM(D.TSA2ClofazimineResult))
                 IN ('Resistant', 'Resistente')
                THEN 'Resistente'

            WHEN NULLIF
                 (
                     LTRIM(RTRIM(D.TSA2ClofazimineResult)),
                     ''
                 ) IS NULL
                THEN 'Não realizado'

            ELSE LTRIM(RTRIM(D.TSA2ClofazimineResult))

        END AS TSA2_Clofazimina

    FROM Dados AS D
)


SELECT

      /* =====================================================
         IDENTIFICAÇÃO DO PACIENTE
         ===================================================== */

      N.CodigoDISA
          AS [Código do DISA]


    , LTRIM
      (
          RTRIM
          (
              ISNULL(N.PatientFirstName, '')
              +
              CASE
                  WHEN N.PatientFirstName IS NOT NULL
                   AND N.PatientSurname IS NOT NULL
                      THEN ' '
                  ELSE ''
              END
              +
              ISNULL(N.PatientSurname, '')
          )
      ) AS [Nome do Paciente]


    , N.AgeInYears
          AS [Idade]


    , N.HL7SexCode
          AS [Sexo]



      /* =====================================================
         UNIDADE / GEOGRAFIA
         ===================================================== */

    , N.RequestingFacilityName
          AS [US de Colheita]


    , CASE

          WHEN N.RequestingDistrictName IS NULL
              THEN N.RequestingProvinceName

          WHEN N.RequestingProvinceName IS NULL
              THEN N.RequestingDistrictName

          ELSE
              N.RequestingDistrictName
              + ' / '
              + N.RequestingProvinceName

      END AS [Distrito / Província]



      /* =====================================================
         AMOSTRA
         ===================================================== */

    , N.LIMSSpecimenSourceDesc
          AS [Tipo de Amostra]



      /* =====================================================
         DATAS
         ===================================================== */

    , CAST(N.SpecimenDateTime AS DATE)
          AS [Data de Colheita]


    , CAST(N.RegisteredDateTime AS DATE)
          AS [Data de Registo no Lab]


    , CAST(N.AnalysisDateTime AS DATE)
          AS [Data de Processamento]


    , CAST(N.AuthorisedDateTime AS DATE)
          AS [Data de Validação]



      /* =====================================================
         TAT

         Registo no laboratório → validação
         ===================================================== */

    , CASE

          WHEN N.RegisteredDateTime IS NOT NULL
           AND N.AuthorisedDateTime IS NOT NULL
           AND N.AuthorisedDateTime >= N.RegisteredDateTime

              THEN DATEDIFF
                   (
                       DAY,
                       N.RegisteredDateTime,
                       N.AuthorisedDateTime
                   )

          ELSE NULL

      END AS [TAT Registo -> Validação (dias)]



      /* =====================================================
         CULTURA
         ===================================================== */

    , N.MGIT_Final
          AS [Resultado Final Cultura Líquida (MGIT)]


    , N.LJ_Final
          AS [Resultado Final Cultura Sólida (LJ)]



      /* =====================================================
         LPA 1
         ===================================================== */

    , N.LPA1_Isoniazida
          AS [LPA1 - Isoniazida]


    , N.LPA1_Rifampicina
          AS [LPA1 - Rifampicina]



      /* =====================================================
         LPA 2

         Uma única coluna, listando os resultados disponíveis.
         ===================================================== */

    , CASE

          WHEN
                 NULLIF(LTRIM(RTRIM(N.LPA2FluoroquinoloneResult)), '') IS NULL
             AND NULLIF(LTRIM(RTRIM(N.LPA2KACResult)), '') IS NULL
             AND NULLIF(LTRIM(RTRIM(N.LPA2KACVResult)), '') IS NULL
             AND NULLIF(LTRIM(RTRIM(N.LPA2KCVResult)), '') IS NULL
             AND NULLIF(LTRIM(RTRIM(N.LPA2LowLevelKanamycinResult)), '') IS NULL

              THEN 'Não realizado'


          ELSE

              STUFF
              (
                    CASE

                        WHEN NULLIF
                             (
                                 LTRIM
                                 (
                                     RTRIM
                                     (
                                         N.LPA2FluoroquinoloneResult
                                     )
                                 ),
                                 ''
                             ) IS NOT NULL

                            THEN
                                '; FLQ: '
                                +
                                CASE
                                    WHEN LTRIM(RTRIM(N.LPA2FluoroquinoloneResult))
                                         IN ('Susceptible','Sensível','Sensivel')
                                        THEN 'Sensível'

                                    WHEN LTRIM(RTRIM(N.LPA2FluoroquinoloneResult))
                                         IN ('Resistant','Resistente')
                                        THEN 'Resistente'

                                    ELSE LTRIM(RTRIM(N.LPA2FluoroquinoloneResult))
                                END

                        ELSE ''

                    END


                  +

                    CASE

                        WHEN NULLIF
                             (
                                 LTRIM(RTRIM(N.LPA2KACResult)),
                                 ''
                             ) IS NOT NULL

                            THEN
                                '; KAN/AMK/CAP: '
                                +
                                CASE
                                    WHEN LTRIM(RTRIM(N.LPA2KACResult))
                                         IN ('Susceptible','Sensível','Sensivel')
                                        THEN 'Sensível'

                                    WHEN LTRIM(RTRIM(N.LPA2KACResult))
                                         IN ('Resistant','Resistente')
                                        THEN 'Resistente'

                                    ELSE LTRIM(RTRIM(N.LPA2KACResult))
                                END

                        ELSE ''

                    END


                  +

                    CASE

                        WHEN NULLIF
                             (
                                 LTRIM(RTRIM(N.LPA2LowLevelKanamycinResult)),
                                 ''
                             ) IS NOT NULL

                            THEN
                                '; KAN baixo nível: '
                                +
                                CASE
                                    WHEN LTRIM(RTRIM(N.LPA2LowLevelKanamycinResult))
                                         IN ('Susceptible','Sensível','Sensivel')
                                        THEN 'Sensível'

                                    WHEN LTRIM(RTRIM(N.LPA2LowLevelKanamycinResult))
                                         IN ('Resistant','Resistente')
                                        THEN 'Resistente'

                                    ELSE LTRIM(RTRIM(N.LPA2LowLevelKanamycinResult))
                                END

                        ELSE ''

                    END,

                  1,
                  2,
                  ''
              )

      END AS [LPA2 - Drogas]



      /* =====================================================
         LPA CM
         ===================================================== */

    , N.LPACM_Identificacao
          AS [LPA CM - Identificação]



      /* =====================================================
         TSA 1
         ===================================================== */

    , N.TSA1_Streptomicina
          AS [TSA1 - Streptomicina]

    , N.TSA1_Isoniazida
          AS [TSA1 - Isoniazida]

    , N.TSA1_Rifampicina
          AS [TSA1 - Rifampicina]

    , N.TSA1_Etambutol
          AS [TSA1 - Etambutol]



      /* =====================================================
         TSA 2
         ===================================================== */

    , N.TSA2_Bedaquilina
          AS [TSA2 - Bedaquilina]

    , N.TSA2_Linezolida
          AS [TSA2 - Linezolida]

    , N.TSA2_Moxifloxacina
          AS [TSA2 - Moxifloxacina]


    /*
        TBCultura não possui campo de Delamanida.
        Não devemos inventar o resultado.
    */
    , CAST(NULL AS VARCHAR(20))
          AS [TSA2 - Delamanida]


    , N.TSA2_Clofazimina
          AS [TSA2 - Clofazimina]



      /* =====================================================
         CLASSIFICAÇÃO DA RESISTÊNCIA

         Ordem:
             XR
             Pré-XR
             RR/MR
             Sensível
             N/A
         ===================================================== */

    , CASE

          /* -------------------------------------------------
             XR

             Rifampicina resistente
             +
             Fluoroquinolona resistente
             +
             Bedaquilina ou Linezolida resistente
             ------------------------------------------------- */

          WHEN
          (
                 N.LPA1_Rifampicina = 'Resistente'
              OR N.TSA1_Rifampicina = 'Resistente'
          )

          AND
          (
                 LTRIM(RTRIM(N.LPA2FluoroquinoloneResult))
                        IN ('Resistant','Resistente')

              OR N.TSA2_Moxifloxacina = 'Resistente'

              OR LTRIM(RTRIM(N.TSA2LevofloxacinResult))
                        IN ('Resistant','Resistente')

              OR LTRIM(RTRIM(N.TSA2OfloxacinResult))
                        IN ('Resistant','Resistente')
          )

          AND
          (
                 N.TSA2_Bedaquilina = 'Resistente'
              OR N.TSA2_Linezolida = 'Resistente'
          )

              THEN 'XR'



          /* -------------------------------------------------
             PRÉ-XR

             Rifampicina resistente
             +
             Fluoroquinolona resistente
             ------------------------------------------------- */

          WHEN
          (
                 N.LPA1_Rifampicina = 'Resistente'
              OR N.TSA1_Rifampicina = 'Resistente'
          )

          AND
          (
                 LTRIM(RTRIM(N.LPA2FluoroquinoloneResult))
                        IN ('Resistant','Resistente')

              OR N.TSA2_Moxifloxacina = 'Resistente'

              OR LTRIM(RTRIM(N.TSA2LevofloxacinResult))
                        IN ('Resistant','Resistente')

              OR LTRIM(RTRIM(N.TSA2OfloxacinResult))
                        IN ('Resistant','Resistente')
          )

              THEN 'Pré-XR'



          /* -------------------------------------------------
             RR/MR
             ------------------------------------------------- */

          WHEN
               N.LPA1_Rifampicina = 'Resistente'
            OR N.TSA1_Rifampicina = 'Resistente'

              THEN 'RR/MR'



          /* -------------------------------------------------
             SENSÍVEL

             Existe pelo menos um teste de susceptibilidade
             realizado e nenhum resultado resistente.
             ------------------------------------------------- */

          WHEN
          (
                 N.LPA1_Isoniazida <> 'Não realizado'
              OR N.LPA1_Rifampicina <> 'Não realizado'

              OR N.TSA1_Streptomicina <> 'Não realizado'
              OR N.TSA1_Isoniazida <> 'Não realizado'
              OR N.TSA1_Rifampicina <> 'Não realizado'
              OR N.TSA1_Etambutol <> 'Não realizado'

              OR N.TSA2_Bedaquilina <> 'Não realizado'
              OR N.TSA2_Linezolida <> 'Não realizado'
              OR N.TSA2_Moxifloxacina <> 'Não realizado'
              OR N.TSA2_Clofazimina <> 'Não realizado'
          )

              THEN 'Sensível'


          ELSE 'N/A'

      END AS [Classificação da Resistência]



      /* =====================================================
         LINK DO LAUDO
         ===================================================== */

    , CASE
          WHEN N.CodigoDISA IS NOT NULL
              THEN @BaseLaudoURL + N.CodigoDISA
          ELSE NULL
      END AS [Link do Laudo]


FROM Normalizado AS N


ORDER BY
      N.AnalysisDateTime
    , N.CodigoDISA;
