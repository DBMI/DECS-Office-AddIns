SET TRANSACTION ISOLATION LEVEL SNAPSHOT

DECLARE @param0031 AS nvarchar ( MAX )  = '["Alcoholic cirrhosis of liver without ascites (CMS-HCC)( ICD-10-CM: K70.30 )","Alcoholic cirrhosis of liver with ascites (CMS-HCC)( ICD-10-CM: K70.31 )","Unspecified cirrhosis of liver (CMS-HCC)( ICD-10-CM: K74.60 )","Other cirrhosis of liver (CMS-HCC)( ICD-10-CM: K74.69 )","Alcoholic hepatitis with ascites( ICD-10-CM: K70.11 )","Alcoholic hepatitis without ascites( ICD-10-CM: K70.10 )","Autoimmune hepatitis (CMS-HCC)( ICD-10-CM: K75.4 )","Liver cell carcinoma (CMS-HCC)( ICD-10-CM: C22.0 )","Angiosarcoma of liver (CMS-HCC)( ICD-10-CM: C22.3 )","Wilson''s disease( ICD-10-CM: E83.01 )","Toxic liver disease with chronic persistent hepatitis( ICD-10-CM: K71.3 )","Toxic liver disease with acute hepatitis( ICD-10-CM: K71.2 )","Other specified diseases of liver( ICD-10-CM: K76.89 )","Liver transplant infection (CMS-HCC)( ICD-10-CM: T86.43 )","Acute hepatitis E( ICD-10-CM: B17.2 )","Carrier of viral hepatitis b( ICD-10-CM: Z22.51 )","Hepatitis a with hepatic coma( ICD-10-CM: B15.0 )","Hepatitis a without hepatic coma( ICD-10-CM: B15.9 )","Human immunodeficiency virus (HIV) disease (CMS-HCC)( ICD-10-CM: B20 )","Asymptomatic human immunodeficiency virus (hiv) infection status (CMS-HCC)( ICD-10-CM: Z21 )","Tuberculosis of lung( ICD-10-CM: A15.0 )","Tuberculosis of cervix( ICD-10-CM: A18.16 )","Tuberculosis of spine( ICD-10-CM: A18.01 )","Tuberculosis of heart( ICD-10-CM: A18.84 )","Tuberculosis of spleen( ICD-10-CM: A18.85 )","Tuberculosis of bladder( ICD-10-CM: A18.12 )","Tuberculosis of prostate( ICD-10-CM: A18.14 )","Tuberculosis of other bones( ICD-10-CM: A18.03 )","Tuberculosis of other sites( ICD-10-CM: A18.89 )","Tuberculosis of adrenal glands( ICD-10-CM: A18.7 )","Tuberculosis of thyroid gland( ICD-10-CM: A18.81 )","Tuberculosis of (inner) (middle) ear( ICD-10-CM: A18.6 )","Tuberculosis of kidney and ureter( ICD-10-CM: A18.11 )","Tuberculosis of eye, unspecified( ICD-10-CM: A18.50 )","Transplanted organ and tissue status, unspecified( ICD-10-CM: Z94.9 )","Elevated urine levels of drugs, medicaments and biological substances( ICD-10-CM: R82.5 )","Congenital renal failure( ICD-10-CM: P96.0 )","Kidney transplant failure( ICD-10-CM: T86.12 )","Kidney transplant rejection( ICD-10-CM: T86.11 )","Heart transplant rejection (CMS-HCC)( ICD-10-CM: T86.21 )"]'

DECLARE @param0032 AS nvarchar ( MAX )  = '["Fibrosis and cirrhosis of liver( ICD-10-CM: K74.* )","Acute hepatitis A( ICD-10-CM: B15.* )","Acute hepatitis B( ICD-10-CM: B16.* )","Asymptomatic human immunodeficiency virus (hiv) infection status (CMS-HCC)( ICD-10-CM: Z21.* )","Tuberculosis of other organs( ICD-10-CM: A18.* )","Tuberculosis of nervous system( ICD-10-CM: A17.* )","Transplanted organ and tissue status( ICD-10-CM: Z94.* )"]'

DECLARE @param0033 AS nvarchar ( MAX )  = '["1150000017","1150000066","102150","102149","102148","102063","2100008010","30218010","101217","118496","101217"]'

DROP TABLE IF EXISTS #JsonTableparam0031

SELECT
 [Value] INTO #JsonTableparam0031 
FROM
 OPENJSON ( @param0031 )  WITH  ( [Value] nvarchar ( 850 )  '$' ) 


DROP TABLE IF EXISTS #JsonTableparam0032

SELECT
 [Value] INTO #JsonTableparam0032 
FROM
 OPENJSON ( @param0032 )  WITH  ( [Value] nvarchar ( 850 )  '$' ) 


DROP TABLE IF EXISTS #JsonTableparam0033

SELECT
 [Value] INTO #JsonTableparam0033 
FROM
 OPENJSON ( @param0033 )  WITH  ( [Value] nvarchar ( 50 )  '$' ) 


DROP TABLE IF EXISTS #GrouperTableparam0086

SELECT
 * INTO #GrouperTableparam0086 
FROM
  ( 
    SELECT
     * 
    FROM
      ( 
        SELECT
         DISTINCT table0.MedicationKey AS basekey, NULL AS grouper_id0, table0.ValueSetEpicId AS grouper_id1, NULL AS grouper_name0, CASE 
        WHEN table0.DisplayName = '*Unspecified' THEN table0.Name 
        ELSE table0.DisplayName END AS grouper_name1 
        FROM
         dbo.MedicationSetDim AS table0 
        WHERE
          (  ( table0.Trusted = 1 
                AND table0.[Type] = 'Simple Generic Name' )  
            AND  ( table0.ValueSetEpicId IS NOT NULL )  )  )  table3 
    WHERE
      (   (  table3.grouper_id1 = '1300001020'  )   OR  (  table3.grouper_id1 = '2150001200'  )   OR  (  table3.grouper_id1 = '2760706010'  )   OR  (  table3.grouper_id1 = '2799780260'  )   OR  (  table3.grouper_id1 = '2728006000'  )   OR  (  table3.grouper_id1 = '2799500270'  )   OR  (  table3.grouper_id1 = '9025054200'  )   )  UNION ALL 
    SELECT
     * 
    FROM
      ( 
        SELECT
         DISTINCT table0.MedicationKey AS basekey, CONCAT ( CONCAT ( table0.ValueSetEpicId,N'-' ) ,table0.Type )  AS grouper_id0, NULL AS grouper_id1, CASE 
        WHEN table0.DisplayName = '*Unspecified' THEN table0.Name 
        ELSE table0.DisplayName END AS grouper_name0, NULL AS grouper_name1 
        FROM
         dbo.MedicationSetDim AS table0 
        WHERE
          (  ( table0.Trusted = 1 
                AND table0.[Type] <> 'Simple Generic Name' )  
            AND  ( CONCAT ( CONCAT ( table0.ValueSetEpicId,N'-' ) ,table0.Type )  IS NOT NULL )  )  )  table3 
    WHERE
      (   (  table3.grouper_id0 = '3800-Pharmaceutical Subclass'  )   OR  (  table3.grouper_id0 = '111680-Epic Medication Grouper'  )   )  )  table3

DROP TABLE IF EXISTS #GrouperTableparam0087

SELECT
 * INTO #GrouperTableparam0087 
FROM
  ( 
    SELECT
     * 
    FROM
      ( 
        SELECT
         DISTINCT table0.DiagnosisKey AS basekey, NULL AS grouper_id0, table0.NameAndCode AS grouper_id1, NULL AS grouper_id2, NULL AS grouper_name0, table0.NameAndCode AS grouper_name1, NULL AS grouper_name2 
        FROM
         dbo.DiagnosisTerminologyDim AS table0 
        WHERE
          (  ( table0.[Type] = 'ICD-10-CM' )  
            AND  ( table0.NameAndCode IS NOT NULL )  )  )  table3 
    WHERE
     EXISTS ( 
        SELECT
         1 
        FROM
         #JsonTableparam0031 jsonTableparam0031 
        WHERE
         jsonTableparam0031.[Value] = table3.grouper_id1 )   UNION ALL 
    SELECT
     * 
    FROM
      ( 
        SELECT
         DISTINCT table0.DiagnosisKey AS basekey, NULL AS grouper_id0, NULL AS grouper_id1, table0.GroupedNameAndCode AS grouper_id2, NULL AS grouper_name0, NULL AS grouper_name1, table0.GroupedNameAndCode AS grouper_name2 
        FROM
         dbo.DiagnosisTerminologyDim AS table0 
        WHERE
          (  ( table0.[Type] = 'ICD-10-CM' )  
            AND  ( table0.GroupedNameAndCode IS NOT NULL )  )  )  table3 
    WHERE
     EXISTS ( 
        SELECT
         1 
        FROM
         #JsonTableparam0032 jsonTableparam0032 
        WHERE
         jsonTableparam0032.[Value] = table3.grouper_id2 )   UNION ALL 
    SELECT
     * 
    FROM
      ( 
        SELECT
         DISTINCT table0.DiagnosisKey AS basekey, table0.ValueSetEpicId AS grouper_id0, NULL AS grouper_id1, NULL AS grouper_id2, CASE 
        WHEN table0.DisplayName = '*Unspecified' THEN table0.Name 
        ELSE table0.DisplayName END AS grouper_name0, NULL AS grouper_name1, NULL AS grouper_name2 
        FROM
         dbo.DiagnosisSetDim AS table0 
        WHERE
          (  ( table0.Trusted = 1 )  
            AND  ( table0.ValueSetEpicId IS NOT NULL )  )  )  table3 
    WHERE
     EXISTS ( 
        SELECT
         1 
        FROM
         #JsonTableparam0033 jsonTableparam0033 
        WHERE
         jsonTableparam0033.[Value] = table3.grouper_id0 )    )  table3

DROP TABLE IF EXISTS #GrouperTableparam0088

SELECT
 * INTO #GrouperTableparam0088 
FROM
  ( 
    SELECT
     * 
    FROM
      ( 
        SELECT
         DISTINCT table0.DiagnosisKey AS basekey, NULL AS grouper_id0, table0.NameAndCode AS grouper_id1, NULL AS grouper_id2, NULL AS grouper_name0, table0.NameAndCode AS grouper_name1, NULL AS grouper_name2 
        FROM
         dbo.DiagnosisTerminologyDim AS table0 
        WHERE
          (  ( table0.[Type] = 'ICD-10-CM' )  
            AND  ( table0.NameAndCode IS NOT NULL )  )  )  table3 
    WHERE
      (   (  table3.grouper_id1 = 'Alcoholic fatty liver( ICD-10-CM: K70.0 )'  )   OR  (  table3.grouper_id1 = 'Alcoholic liver disease, unspecified (CMS-HCC)( ICD-10-CM: K70.9 )'  )   OR  (  table3.grouper_id1 = 'Alcohol use, unspecified, uncomplicated( ICD-10-CM: F10.90 )'  )   OR  (  table3.grouper_id1 = 'Alcohol dependence, uncomplicated (CMS-HCC)( ICD-10-CM: F10.20 )'  )   OR  (  table3.grouper_id1 = 'Alcohol dependence with withdrawal delirium (CMS-HCC)( ICD-10-CM: F10.231 )'  )   OR  (  table3.grouper_id1 = 'Alcohol dependence with intoxication delirium (CMS-HCC)( ICD-10-CM: F10.221 )'  )   OR  (  table3.grouper_id1 = 'Alcohol dependence with withdrawal, unspecified (CMS-HCC)( ICD-10-CM: F10.239 )'  )   OR  (  table3.grouper_id1 = 'Alcohol dependence with intoxication, uncomplicated (CMS-HCC)( ICD-10-CM: F10.220 )'  )   OR  (  table3.grouper_id1 = 'Alcohol dependence with intoxication, unspecified (CMS-HCC)( ICD-10-CM: F10.229 )'  )   OR  (  table3.grouper_id1 = 'Alcohol dependence with other alcohol-induced disorder (CMS-HCC)( ICD-10-CM: F10.288 )'  )   OR  (  table3.grouper_id1 = 'Alcohol dependence with withdrawal, uncomplicated (CMS-HCC)( ICD-10-CM: F10.230 )'  )   OR  (  table3.grouper_id1 = 'Alcohol use, unspecified with intoxication delirium (CMS-HCC)( ICD-10-CM: F10.921 )'  )   OR  (  table3.grouper_id1 = 'Alcohol use, unspecified with withdrawal delirium (CMS-HCC)( ICD-10-CM: F10.931 )'  )   OR  (  table3.grouper_id1 = 'Alcohol use, unspecified with withdrawal, unspecified (CMS-HCC)( ICD-10-CM: F10.939 )'  )   OR  (  table3.grouper_id1 = 'Liver disease, unspecified( ICD-10-CM: K76.9 )'  )   OR  (  table3.grouper_id1 = 'Hepatic fibrosis( ICD-10-CM: K74.0 )'  )   OR  (  table3.grouper_id1 = 'Hepatic fibrosis, unspecified( ICD-10-CM: K74.00 )'  )   OR  (  table3.grouper_id1 = 'Hepatic fibrosis, early fibrosis( ICD-10-CM: K74.01 )'  )   OR  (  table3.grouper_id1 = 'Alcohol dependence, in remission (CMS-HCC)( ICD-10-CM: F10.21 )'  )   OR  (  table3.grouper_id1 = 'Alcohol abuse with withdrawal delirium (CMS-HCC)( ICD-10-CM: F10.131 )'  )   OR  (  table3.grouper_id1 = 'Alcohol abuse with withdrawal, unspecified (CMS-HCC)( ICD-10-CM: F10.139 )'  )   OR  (  table3.grouper_id1 = 'Alcohol abuse with withdrawal, uncomplicated (CMS-HCC)( ICD-10-CM: F10.130 )'  )   OR  (  table3.grouper_id1 = 'Alcohol abuse with alcohol-induced mood disorder (CMS-HCC)( ICD-10-CM: F10.14 )'  )   OR  (  table3.grouper_id1 = 'Alcohol abuse counseling and surveillance of alcoholic( ICD-10-CM: Z71.41 )'  )   OR  (  table3.grouper_id1 = 'Alcohol abuse with intoxication, unspecified (CMS-HCC)( ICD-10-CM: F10.129 )'  )   OR  (  table3.grouper_id1 = 'Alcohol abuse, in remission( ICD-10-CM: F10.11 )'  )   OR  (  table3.grouper_id1 = 'Alcohol abuse, uncomplicated( ICD-10-CM: F10.10 )'  )   OR  (  table3.grouper_id1 = 'Alcohol use, unspecified, in remission( ICD-10-CM: F10.91 )'  )   OR  (  table3.grouper_id1 = 'Alcohol abuse with intoxication delirium (CMS-HCC)( ICD-10-CM: F10.121 )'  )   OR  (  table3.grouper_id1 = 'Abnormal results of liver function studies( ICD-10-CM: R94.5 )'  )   OR  (  table3.grouper_id1 = 'Elevation of levels of liver transaminase levels( ICD-10-CM: R74.01 )'  )   )  UNION ALL 
    SELECT
     * 
    FROM
      ( 
        SELECT
         DISTINCT table0.DiagnosisKey AS basekey, NULL AS grouper_id0, NULL AS grouper_id1, table0.GroupedNameAndCode AS grouper_id2, NULL AS grouper_name0, NULL AS grouper_name1, table0.GroupedNameAndCode AS grouper_name2 
        FROM
         dbo.DiagnosisTerminologyDim AS table0 
        WHERE
          (  ( table0.[Type] = 'ICD-10-CM' )  
            AND  ( table0.GroupedNameAndCode IS NOT NULL )  )  )  table3 
    WHERE
      (   (  table3.grouper_id2 = 'Alcoholic liver disease( ICD-10-CM: K70.* )'  )   OR  (  table3.grouper_id2 = 'Alcohol related disorders( ICD-10-CM: F10.* )'  )   )  UNION ALL 
    SELECT
     * 
    FROM
      ( 
        SELECT
         DISTINCT table0.DiagnosisKey AS basekey, table0.ValueSetEpicId AS grouper_id0, NULL AS grouper_id1, NULL AS grouper_id2, CASE 
        WHEN table0.DisplayName = '*Unspecified' THEN table0.Name 
        ELSE table0.DisplayName END AS grouper_name0, NULL AS grouper_name1, NULL AS grouper_name2 
        FROM
         dbo.DiagnosisSetDim AS table0 
        WHERE
          (  ( table0.Trusted = 1 )  
            AND  ( table0.ValueSetEpicId IS NOT NULL )  )  )  table3 
    WHERE
      (  table3.grouper_id0 = '102159'  )   )  table3


DROP TABLE IF EXISTS #resultSet

SELECT
 * INTO #resultSet 
FROM
 ( 

    SELECT
     COUNT_BIG ( DurableKey )  AS count0 , COUNT_BIG ( DurableKey )  AS count1  

    FROM
      ( 

        SELECT
         subq0.DurableKey, NULL AS count0, NULL AS count1 
        FROM
          (  
            SELECT
             table0.DurableKey AS DurableKey, NULL AS count0 , NULL AS count1  
            FROM
             dbo.PatientDim AS table0 
            WHERE
              (  (  ( table0.IsValid = 1 )  
                    AND  ( table0.IsHistoricalPatient = 0 )  )  
                AND  ( table0.IsCurrent = 1 )  
                AND EXISTS  ( 
                    SELECT
                     1 
                    FROM
                     Epic.SdSocialHistoryFact AS table1 
                    WHERE
                      (   ( table1.NumericData >= 1 )   
                        AND  ( table1.IsCurrent = 1 )   
                        AND  ( table1.FilterId = '37' )   
                        AND  (  table0.PatientEpicId = table1.PatientEpicId  )  )  )   
                AND NOT  (  EXISTS  ( 
                        SELECT
                         1 
                        FROM
                         dbo.ProcedureEventFact AS table1 
                        WHERE
                          (   (   (  table1.ProcedureDurableKey = '32171'  )   OR  (  table1.ProcedureDurableKey = '31775'  )   OR  (  table1.ProcedureDurableKey = '238280'  )   OR  (  table1.ProcedureDurableKey = '31246'  )   OR  (  table1.ProcedureDurableKey = '31247'  )   OR  (  table1.ProcedureDurableKey = '18863'  )   OR  (  table1.ProcedureDurableKey = '37939'  )   OR  (  table1.ProcedureDurableKey = '5766'  )   OR  (  table1.ProcedureDurableKey = '31774'  )   OR  (  table1.ProcedureDurableKey = '31020'  )   OR  (  table1.ProcedureDurableKey = '32916'  )   OR  (  table1.ProcedureDurableKey = '17585'  )   OR  (  table1.ProcedureDurableKey = '32801'  )   OR  (  table1.ProcedureDurableKey = '32129'  )   OR  (  table1.ProcedureDurableKey = '140944'  )   OR  (  table1.ProcedureDurableKey = '140945'  )   OR  (  table1.ProcedureDurableKey = '181724'  )   OR  (  table1.ProcedureDurableKey = '140947'  )   OR  (  table1.ProcedureDurableKey = '32172'  )   OR  (  table1.ProcedureDurableKey = '32173'  )   )   
                            AND  ( table1.SourceKey IN  ( 1, 3, 4, 5, 7, 8, 9, 11, 12, 13, 14, 15, 16, 17, 19, 20, 21, 22, 23, 362, -1 )  )  
                            AND  ( table1.[Type] = 'Surgical History Procedure' )   
                            AND  (  table0.DurableKey = table1.PatientDurableKey  )  )  )   )  
                AND NOT  (  EXISTS  ( 
                        SELECT
                         1 
                        FROM
                         dbo.MedicationEventFact AS table1 
                        INNER JOIN #GrouperTableparam0086 AS table3 ON table1.MedicationKey = table3.basekey  
                        WHERE
                          (  (  (  ( table1.StartInstant < '11/21/2023 12:00:00 AM' )  
                                    AND NOT  ( table1.StartInstant IS NULL )  )  
                                AND  (  ( table1.EndInstant >= '11/20/2022 12:00:00 AM' )  OR  ( table1.EndInstant IS NULL )  )  )  
                            AND  ( table1.SourceKey IN  ( 1, 3, 4, 5, 7, 8, 9, 11, 12, 13, 14, 15, 16, 17, 19, 20, 21, 22, 23, 362, -1 )  )  
                            AND  ( table1.[Type] IN  ( 'Medication Order with Administration', 'Medication Order', 'Inpatient', 'Outpatient', 'Historical', 'Merged Medication Timeline', 'Administration'  )  )   
                            AND  (  table0.DurableKey = table1.PatientDurableKey  )  )  )   )  
                AND NOT  (  EXISTS  ( 
                        SELECT
                         1 
                        FROM
                         dbo.DiagnosisEventFact AS table1 
                        INNER JOIN #GrouperTableparam0087 AS table3 ON table1.DiagnosisKey = table3.basekey  
                        WHERE
                          (  (  (  ( table1.StartDateKey < '20231121' )  
                                    AND NOT  ( table1.StartDateKey < 0 )  )  
                                AND  (  ( table1.EndDateKey >= '20221120' )  OR  ( table1.EndDateKey < 0 )  )  )  
                            AND  ( table1.SourceKey IN  ( 1, 3, 4, 5, 7, 8, 9, 11, 12, 13, 14, 15, 16, 17, 19, 20, 21, 22, 23, 362, -1 )  )  
                            AND  ( table1.[Type] IN  ( 'Encounter Diagnosis', 'Billing Diagnosis', 'Problem List', 'Hospital Problem', 'Admitting Diagnosis', 'Discharge Diagnosis' )  )   
                            AND  (  table0.DurableKey = table1.PatientDurableKey  )  )  )   )  
                AND EXISTS  ( 
                    SELECT
                     1 
                    FROM
                     dbo.AddressDim AS table1 
                    WHERE
                      (   (  CONCAT ( table1.County, CASE 
                                WHEN table1.StateOrProvinceAbbreviation<>'*Unspecified' THEN CONCAT ( ', ', table1.StateOrProvinceAbbreviation )  
                                ELSE '' END )  = 'SAN DIEGO, CA'  )   
                        AND  ( table1.County NOT IN  ( '*Unspecified','*Unknown','*Deleted','*Not Applicable' )  )   
                        AND  (  table0.AddressKey = table1.AddressKey  )  )  )   
                AND  (  CASE 
                    WHEN table0.Status IN  ( '*Not Applicable', '*Unknown', '*Unspecified', '*Deleted', '' )  THEN 'Alive' 
                    ELSE table0.Status END = 'Alive'  )   
                AND EXISTS  ( 
                    SELECT
                     1 
                    FROM
                     Epic.SdVitalsFact AS table1 
                    WHERE
                      (   ( table1.NumericData <= 45 )   
                        AND  ( table1.NumericData >= 20 )   
                        AND  (  (  ( table1.EffectiveStartDate < '11/21/2023 12:00:00 AM' )  OR  ( table1.EffectiveStartDate IS NULL )  )  
                            AND  (  ( table1.EffectiveEndDate >= '11/20/2022 12:00:00 AM' )  OR  ( table1.EffectiveEndDate IS NULL )  )  )  
                        AND  ( table1.IsCurrent = 1 )   
                        AND  ( table1.FilterId = '20' )   
                        AND  (  table0.PatientEpicId = table1.PatientEpicId  )  )  )   
                AND  ( table0.AgeInYears >= 21 )   
                AND EXISTS  ( 
                    SELECT
                     1 
                    FROM
                     dbo.DiagnosisEventFact AS table1 
                    INNER JOIN #GrouperTableparam0088 AS table3 ON table1.DiagnosisKey = table3.basekey  
                    WHERE
                      (  (  (  ( table1.StartDateKey < '20231121' )  
                                AND NOT  ( table1.StartDateKey < 0 )  )  
                            AND  (  ( table1.EndDateKey >= '20221120' )  OR  ( table1.EndDateKey < 0 )  )  )  
                        AND  ( table1.SourceKey IN  ( 1, 3, 4, 5, 7, 8, 9, 11, 12, 13, 14, 15, 16, 17, 19, 20, 21, 22, 23, 362, -1 )  )  
                        AND  ( table1.[Type] IN  ( 'Encounter Diagnosis', 'Billing Diagnosis', 'Problem List', 'Hospital Problem', 'Admitting Diagnosis', 'Discharge Diagnosis' )  )   
                        AND  (  table0.DurableKey = table1.PatientDurableKey  )  )  )   
                AND EXISTS  ( 
                    SELECT
                     1 
                    FROM
                     dbo.LabComponentResultFact AS table1 
                    WHERE
                      (   ( table1.NumericValue >= 3.7001 )   
                        AND  ( table1.Unit = table1.Unit )   
                        AND  (  (  ( table1.PrioritizedDateKey < '20231121' )  
                                AND NOT  ( table1.PrioritizedDateKey < 0 )  )  
                            AND  (  ( table1.PrioritizedDateKey >= '20221120' )  OR  ( table1.PrioritizedDateKey < 0 )  )  )  
                        AND  (  table1.LabComponentKey IN  ( '1', '242', '8102', '13633', '26700', '27885' )   )   
                        AND  ( table1.SourceKey IN  ( 1, 3, 4, 5, 7, 8, 9, 11, 12, 13, 14, 15, 16, 17, 19, 20, 21, 22, 23, 362, -1 )  )  
                        AND  ( table1.IsBlankOrUnsuccessfulAttempt = 0 )   
                        AND  (  table0.DurableKey = table1.PatientDurableKey  )  )  )   
                AND EXISTS  ( 
                    SELECT
                     1 
                    FROM
                     dbo.LabComponentResultFact AS table1 
                    WHERE
                      (   ( table1.NumericValue <= 1.2999 )   
                        AND  ( table1.Unit = table1.Unit )   
                        AND  (  (  ( table1.PrioritizedDateKey < '20231121' )  
                                AND NOT  ( table1.PrioritizedDateKey < 0 )  )  
                            AND  (  ( table1.PrioritizedDateKey >= '20221120' )  OR  ( table1.PrioritizedDateKey < 0 )  )  )  
                        AND  (  table1.LabComponentKey IN  ( '1108', '8342' )   )   
                        AND  ( table1.SourceKey IN  ( 1, 3, 4, 5, 7, 8, 9, 11, 12, 13, 14, 15, 16, 17, 19, 20, 21, 22, 23, 362, -1 )  )  
                        AND  ( table1.IsBlankOrUnsuccessfulAttempt = 0 )   
                        AND  (  table0.DurableKey = table1.PatientDurableKey  )  )  )   
                AND EXISTS  ( 
                    SELECT
                     1 
                    FROM
                     dbo.LabComponentResultFact AS table1 
                    WHERE
                      (   ( table1.NumericValue <= 1.4999 )   
                        AND  ( table1.NumericValue >= 0 )   
                        AND  ( table1.Unit = table1.Unit )   
                        AND  (  (  ( table1.PrioritizedDateKey < '20231121' )  
                                AND NOT  ( table1.PrioritizedDateKey < 0 )  )  
                            AND  (  ( table1.PrioritizedDateKey >= '20221120' )  OR  ( table1.PrioritizedDateKey < 0 )  )  )  
                        AND  (  table1.LabComponentKey IN  ( '231', '7851', '26712', '27877', '35327' )   )   
                        AND  ( table1.SourceKey IN  ( 1, 3, 4, 5, 7, 8, 9, 11, 12, 13, 14, 15, 16, 17, 19, 20, 21, 22, 23, 362, -1 )  )  
                        AND  ( table1.IsBlankOrUnsuccessfulAttempt = 0 )   
                        AND  (  table0.DurableKey = table1.PatientDurableKey  )  )  )   
                AND EXISTS  ( 
                    SELECT
                     1 
                    FROM
                     dbo.PatientServiceAreaMappingFact AS table1 
                    WHERE
                      (   (  table1.ServiceAreaEpicId = '10'  )   
                        AND  (  table0.DurableKey = table1.PatientDurableKey  )  )  )   )   )  subq0 
 )  subq

    WHERE
     
	subq.DurableKey > 0 

 )  AS tempResultSet 
OPTION ( RECOMPILE, USE HINT ( 'FORCE_DEFAULT_CARDINALITY_ESTIMATION', 'DISABLE_OPTIMIZER_ROWGOAL' ) , NO_PERFORMANCE_SPOOL ) 

SELECT
 DISTINCT  mainResultSet.count0 AS count0 ,  mainResultSet.count1 AS count1  
FROM
 #resultSet mainResultSet 

DROP TABLE #resultSet