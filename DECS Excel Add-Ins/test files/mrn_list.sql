USE [REL_CLARITY];

DROP TABLE IF EXISTS #PATIENT_LIST;
CREATE TABLE #PATIENT_LIST (MRN varchar(18))
:setvar path "F:\DECS\<task folder name>"
:r $(path)\mrn_list.sql
