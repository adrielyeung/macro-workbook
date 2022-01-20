# macro-workbook
A collection of VBA Excel macros that I use for utility.

## How to use
Press Alt+F11 on Excel to turn on Visual Basic. Import the macro by right clicking on your excel workbook name and select 'Import', then select the correct ```.bas``` file.

## Description
```FindKeyword.bas```
---------------------
This is a macro used to search for a number of keywords in the config file (with comma-separated keyword values).

For each row, the found keyword(s) in search range will be put into the first empty column (Result Column), as a comma-separated list.

A relevancy score is calculated and placed into the second empty column. The results are sorted according to descending relevancy.

Inputs (via InputBox / MsgBox):
1. Config file path
2. Search Range (e.g. "A1:B2" Range in Excel)
3. Whether to use equal scoring for each keyword, or weighted scoring (weights will be added in descending order of each word in the config file). The final score would be used to sort the results.

All 3 inputs are compulsory.

```MergeEmptyVertical.bas```
----------------------------
This is a macro used to merge empty cells in a column with the nearest non-empty cell above.

Additionally, can output the number of rows merged for each non-empty cell to a specified column (requires the non-empty cell to be on same row, as in below screenshot).

<img src="https://github.com/adrielyeung/macro-workbook/blob/main/img/MergeEmptyVertical_before.png" alt="MergeEmptyVertical_before" width="25%" height="25%">
<img src="https://github.com/adrielyeung/macro-workbook/blob/main/img/MergeEmptyVertical_after.png" alt="MergeEmptyVertical_after" width="25%" height="25%">

```WordUtil.bas```
------------------
```CreateWordDoc()```

Creates a Word document.

```ReplaceTagsWithContent()```

Fill content into a Word template by replacing tags ```<xxx>``` with content.

Example Cover Letter template and config are provided in ```config/``` folder.

Support for paragraph building with multiple sentences in same category (```P<xxx>``` tags in Excel config). This is done by extracting stock phrases from PhraseConfig sheet with a random starting phrase, and inserting the config data into the stock phrases.

## Future developments
