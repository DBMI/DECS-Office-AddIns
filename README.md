# DECS Word Add-Ins ![image info](./DECS%20Word%20Add-Ins/pictures/toolbar_mini.png) 

![Last Commit Date](./DECS%20Word%20Add-Ins/.github/badges/last-commit-badge.svg?dummy=8484744)

Creates a custom buttons in Microsoft Word that allows user to:
* Scan a statement of work file & create SQL code that searches for the ICD-10 codes and names listed in the SoW.
* Turn a list of MRNs into a SQL snippet that imports the list into a query.

## Installation
* Download the `Word Add-Ins` folder from [Sharepoint.](https://ucsdhs.sharepoint.com/:f:/t/ACTRI-BMI-DECSPrivate/EhFYD_9zfX9GsNRN9enCMzABFKg6wmPh13zY_ps2qRJHSg?e=KYFZeG)
* Run `setup.exe`.

## Operation
### Extract ICD codes
Sometimes Statements of Work (SoW) contain lists of medical conditions and ICD-10 codes to be reported on.
Pressing the `Extract ICD` button causes the app to scan the open Word document for lines that look like medical conditions and their associated ICD-10 codes. SQL code is generated that searches the `problem_list` table for the associated codes, as shown here:![image info](./DECS%20Word%20Add-Ins/pictures/ICD%20to%20sql%20basic.png)

Series of ICD codes (such as `M30 - M36`) are automatically expanded into multi-code SQL statements:![image info](./DECS%20Word%20Add-Ins/pictures/series%20expansion%20sql.png)

### Build MRN Import
When researchers provide lists of Medical Record Numbers (MRNs) to be used in a report, those MRNs need to be imported into SQL. Pressing the `Build MRN Import` button converts a list of numbers into SQL code which can be referenced in a query to import these MRNs:![image info](./DECS%20Word%20Add-Ins/pictures/MRN%20list%20to%20sql%20top.png)
Since there is a limit on the number of values (1000) that can be inserted in one statement, the app automatically breaks up the insertion into multiple statements:

![image info](./DECS%20Word%20Add-Ins/pictures/MRN%20list%20break.png)
