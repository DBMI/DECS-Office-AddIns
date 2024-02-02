# DECS Office Add-Ins
![Last Commit Date](./.github/badges/last-commit-badge.svg?dummy=8484744)

Creates custom buttons in Microsoft Excel & Word that allow user to:
### Excel:
![image info](./DECS%20Excel%20Add-Ins/pictures/toolbar.png) 

* Scan `Notes` fields for keywords, creating new columns.
* Turn a list (of MRNs, ICD codes, etc.) into a SQL snippet that imports the column into a query.
* Lookup Social Vulnerability Index ([SVI](https://www.atsdr.cdc.gov/placeandhealth/svi/index.html)) from address or zip code.
* Format a page of results with bold & centered header, NULLs grayed out, etc.
* Convert dates from [MUMPS](https://en.wikipedia.org/wiki/MUMPS) to Excel standard.
### Word: 
![image info](./DECS%20Word%20Add-Ins/pictures/toolbar.png)
* Scan a Scope of Work (SoW) file & create SQL code that searches for the ICD-9/ICD-10 codes and names listed in the SoW.
* Turn a list of MRNs into a SQL snippet that imports the list into a query.
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
### Excel: Extract data from notes
In cases where we want to extract data from free-text columns, we can use the `Notes` tools to define & run extraction rules.
#### Cleaning
We start in the `Cleaning` tab, defining cleaning rules to:
* fix misspellings
* enforce standard naming

These rules are run *before* the data extraction rules.

![image info](./DECS%20Excel%20Add-Ins/pictures/cleaning%20rules.png)
#### Date formats
Using the `DateFormat` tab, we can select which date format we want for output columns.

![image info](./DECS%20Excel%20Add-Ins/pictures/date%20formats.png)
#### Extraction rules
The `Extract` tab lets us define the Regular Expressions that extract data from free text.

![image info](./DECS%20Excel%20Add-Ins/pictures/extraction%20rules.png)

Starting with these free-text notes:

![image info](./DECS%20Excel%20Add-Ins/pictures/notes%20raw.png)

Here's an example of the extracted data:

![image info](./DECS%20Excel%20Add-Ins/pictures/notes%20results.png)

Notice how the original dates--in multiple formats--were automatically converted to a standard date format before extraction.

### Word: Extract ICD codes
Sometimes Statements of Work (SoW) contain lists of medical conditions and ICD-10 codes to be reported on.
Pressing the `Extract ICD` button causes the app to scan the open Word document for lines that look like medical conditions and their associated ICD-10 codes. SQL code is generated that searches the `problem_list` table for the associated codes, as shown here:![image info](./DECS%20Word%20Add-Ins/pictures/ICD%20to%20sql%20basic.png)

Series of ICD codes (such as `M30 - M36`) are automatically expanded into multi-code SQL statements:![image info](./DECS%20Word%20Add-Ins/pictures/series%20expansion%20sql.png)

### Word: Build MRN Import
When researchers provide lists of Medical Record Numbers (MRNs) to be used in a report, those MRNs need to be imported into SQL. Pressing the `Build MRN Import` button converts a list of numbers into SQL code which can be referenced in a query to import these MRNs:![image info](./DECS%20Word%20Add-Ins/pictures/MRN%20list%20to%20sql%20top.png)

Since there is a limit on the number of values (1000) that can be inserted in one statement, the app automatically breaks up the insertion into multiple statements:

![image info](./DECS%20Word%20Add-Ins/pictures/MRN%20list%20break.png)
