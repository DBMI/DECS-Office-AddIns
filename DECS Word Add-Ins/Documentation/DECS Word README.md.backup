# DECS Word Add-Ins

![DECS custom Word toolbar](toolbar.png)

Creates custom buttons in Microsoft Word that allow user to:

### Word: 
* Scan a Scope of Work (SoW) file & create SQL code that searches for the ICD-9/ICD-10 codes and names listed in the SoW.
* Turn a list of MRNs or ICDs into a SQL snippet that imports the list into a query.
* Setup a DECS project using the info in a Scope of Work file:
    - Build the DECS project directory.
    - Initialize the Excel output file, including disclaimer.
    - Initialize the SQL file.
    - Modify a Slicer/Dicer SQL file to include patient consent, etc.
    - Push the SQL file to GitLab.
    - Creates a project folder in Outlook.
    - Draft the completion email.

## Installation
* Download the `Office Add-Ins` folder from [Sharepoint.](https://ucsdhs.sharepoint.com/:f:/t/ACTRI-BMI-DECSPrivate/EhFYD_9zfX9GsNRN9enCMzABFKg6wmPh13zY_ps2qRJHSg?e=KYFZeG)
* Run `setup.exe`.

## Operation

### Word: Extract ICD codes
Sometimes Statements of Work (SoW) contain lists of medical conditions and ICD-10 codes to be reported on.
Pressing the `Extract ICD` button causes the app to scan the open Word document for lines that look like medical conditions and their associated ICD-10 codes. SQL code is generated that searches the `problem_list` table for the associated codes, as shown here:![Example of ICD list translated to SQL](ICD_to_sql_basic.png)

Series of ICD codes (such as `M30 - M36`) are automatically expanded into multi-code SQL statements:![Expanding to an ICD list](series_expansion_sql.png)

### Word: Build List Import
When researchers provide lists of Medical Record Numbers (MRNs) or International Classification of Diseases (ICD) codes to be used in a report, those lists need to be imported into SQL. Pressing the `Import List` button converts a list of numbers into SQL code which can be referenced in a query to import them:![Converting a list to SQL import](MRN_list_to_sql_top.png)

Since there is a limit on the number of values (1000) that can be inserted in one statement, the app automatically breaks up the insertion into multiple statements:

![image info](MRN_list_break.png)
