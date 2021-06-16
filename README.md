# macro-workbook
A collection of VBA Excel macros that I use for utility.

## How to use
Press Alt+F11 on Excel to turn on Visual Basic. Import the macro by right clicking on your excel workbook name and select 'Import', then select the correct ```.bas``` file.

## Description
```FindKeyword.bas```
---------------------
This is a macro used to search for a number of keywords in the config file (with comma-separated keyword values). For each row, the found keyword(s) in search range will be put into a new column (Result Column), as a comma-separated list.

Inputs (via InputBox):
1. Config file path
2. Search Range (e.g. "A1:B2" Range in Excel)
3. Output column (e.g. "K")

All 3 inputs are compulsory.

## Future developments
```FindKeyword.bas```
---------------------
1. Separate each keyword found into different columns, which would allow for counting of keywords and sort by relevancy.
