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
3. Start of output columns (this should be the first free column, multiple columns will be used) (e.g. "K").
4. Whether to use equal scoring for each keyword, or weighted scoring (weights will be added in descending order of each word in the config file). The final score would be used to sort the results.

All 3 inputs are compulsory.

## Future developments
```FindKeyword.bas```
---------------------
1. Automatic detection of free column for output.
