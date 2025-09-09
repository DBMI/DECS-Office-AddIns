USE [REL_CLARITY];

DROP TABLE IF EXISTS #PATIENT_LIST;
CREATE TABLE #PATIENT_LIST (ZIP_CODE varchar(18))
:setvar path "F:\DECS\<task folder name>"
:r $(path)\patient_zip_codes.sql
