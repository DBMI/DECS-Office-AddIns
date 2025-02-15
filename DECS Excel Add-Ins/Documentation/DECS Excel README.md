# DECS Excel Add-Ins

![DECS custom Excel toolbar](toolbar.png) 

Creates custom buttons in Microsoft Excel that allow user to:

* Scan `Notes` fields for keywords, creating new columns.
* Turn a list (of MRNs, ICD codes, etc.) into a SQL snippet that imports the column into a query.
* Lookup Social Vulnerability Index ([SVI](https://www.atsdr.cdc.gov/placeandhealth/svi/index.html)) from address or zip code.
* Format a page of results with bold & centered header, NULLs grayed out, etc.
* Convert dates from [MUMPS](https://en.wikipedia.org/wiki/MUMPS) to Excel standard.

## Installation
* Download the `Office Add-Ins` folder from [Sharepoint.](https://ucsdhs.sharepoint.com/:f:/t/ACTRI-BMI-DECSPrivate/EhFYD_9zfX9GsNRN9enCMzABFKg6wmPh13zY_ps2qRJHSg?e=KYFZeG)
* Run `setup.exe`.

## Operation
### Extract data from notes
In cases where we want to extract data from free-text columns, we can use the `Notes` tools to define & run extraction rules.
#### Cleaning
We start in the `Cleaning` tab, defining cleaning rules to:
* fix misspellings
* enforce standard naming

These rules are run *before* the data extraction rules.

![Example of data cleaning rules](cleaning_rules.png)
#### Date formats
Using the `DateFormat` tab, we can select which date format we want for output columns.

![Example of date format selection](date_formats.png)
#### Extraction rules
The `Extract` tab lets us define the Regular Expressions that extract data from free text.

![Example of data extraction rules](extraction_rules.png)

Starting with these free-text notes:

![Raw patient notes](notes_raw.png)

Here's an example of the extracted data:

![Processed patient notes](notes_results.png)

Notice how the original dates--in multiple formats--were automatically converted to a standard date format before extraction.

## Algorithm Description
### Excel: Lookup Social Vulnerability Index (SVI)
Researchers sometimes ask for the [Social Vulnerability Index](https://www.atsdr.cdc.gov/placeandhealth/svi/index.html) of a patient's neighborhood in order to provide a fuller picture of their [social determinants of health](https://health.gov/healthypeople/priority-areas/social-determinants-health). According to the US Centers for Disease Control's Agency for Toxic Substances and Disease Registry (CDC/ATSDR):

> Social vulnerability refers to the potential negative effects on communities caused by external stresses on human health. Such stresses include natural or human-caused disasters, or disease outbreaks. Reducing social vulnerability can decrease both human suffering and economic loss.

We can provide this information by using the patient's address (preferred) or zip code to look up their [census tract number](https://www.census.gov/programs-surveys/geography/about/glossary.html#par_textimage_13), then retrieving their SVI info from a [document](https://www.atsdr.cdc.gov/placeandhealth/svi/data_documentation_download.html) published by the CDC/ATDSR, which lists the SVI information for each census tract (called a [Federal Information Processing Standard (FIPS)](https://nitaac.nih.gov/resources/frequently-asked-questions/what-fips-code-and-why-do-i-need-one#:~:text=The%20Federal%20Information%20Processing%20Standard,equivalents%20in%20the%20United%20States.) code in the data files).

#### Finding the census tract number:
**Address**
If we have a patient's address, we can lookup their exact census tract number using the US Census Bureau's [online geocoding service](https://geocoding.geo.census.gov/geocoder/Geocoding_Services_API.html). Here's an example of the query & response (the census tract number is here called *GEOID*):

Query:
https://geocoding.geo.census.gov/geocoder/geographies/onelineaddress?address=1600%20PENNSYLVANIA%20AVE%2C%20WASHINGTON%20DC%2020500&benchmark=2020&vintage=2020&format=json

Response:

![JSON geocoding response](json_response.png)

**Zip code**
Using just a zip code is less exact than a full address, as a zip code may contain many census tracts, and tracts may overlap with more than one zip code. However, we can use [*crosswalk*](https://www.huduser.gov/portal/datasets/usps_crosswalk.html) files provided by the US Department of Housing and Urban Development to lookup all the census tracts present in a given zip code. Here's an example from the file ZIP_TRACT_122023.csv (available [here](https://www.huduser.gov/portal/datasets/usps_crosswalk.html)). Notice the large number of census tracts which cross to a typical San Diego County zip code:

![Zip crosswalk file](crosswalk_multiple_tracts.png)

#### Extracting SVI information
The data files published by CDC/ATDSR have one row of social vulnerability data for each census tract. The meaning of each column is explained [here](https://www.atsdr.cdc.gov/placeandhealth/svi/documentation/SVI_documentation_2020.html); we extract the column SPL_THEMES as the SVI score, and RPL_THEMES as the SVI ranking. While both single-state and entire-USA files are available, we've used the entire-USA file to be able to provide data both within & outside of California. If single-state files are used, please review the section *Caveat for SVI State Databases* in [SVI documentation](https://www.atsdr.cdc.gov/placeandhealth/svi/documentation/SVI_documentation_2020.html) for important statistical concerns.

Here's an example of the SVI data from file SVI_2020_US.csv (available [here](https://www.atsdr.cdc.gov/placeandhealth/svi/data_documentation_download.html)) for the census tract 11001980000 returned in the example query above:

![SVI census data](SVI_example.png)

In cases where we're using only zip code information and have multiple census tracts, we return the average SVI score and ranking across all tracts associated with the zip code.


