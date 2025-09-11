## Extract data from notes
In cases where we want to extract data from free-text columns, we can use the `Notes` tools to define & run extraction rules.

![image info](./toolbar.png)

#### Cleaning
We start in the `Cleaning` tab, defining cleaning rules to:
* fix misspellings
* enforce standard naming

These rules are run *before* the data extraction rules.

![image info](./cleaning_rules.png)

#### Date formats
Using the `DateFormat` tab, we can select which date format we want for output columns.

![image info](./date_formats.png)
#### Extraction rules
The `Extract` tab lets us define the Regular Expressions that extract data from free text.

![image info](./extraction_rules.png)

Starting with these free-text notes:

![image info](./notes_raw.png)

Here's an example of the extracted data:

![image info](./notes_results.png)

Notice how the original dates--in multiple formats--were automatically converted to a standard date format before extraction.

[BACK](../../README.md)