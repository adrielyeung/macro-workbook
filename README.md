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
This is a macro used to merge empty cells in each column in a selected range with the nearest non-empty cell above.

Additionally, can output the number of rows merged for each non-empty cell to a specified column (requires the non-empty cell to be on same row, as in column "D" in below screenshot).

Before:

<img src="https://github.com/adrielyeung/macro-workbook/blob/main/img/MergeEmptyVertical_before.png" alt="MergeEmptyVertical_before" width="50%" height="50%">

After:

<img src="https://github.com/adrielyeung/macro-workbook/blob/main/img/MergeEmptyVertical_after.png" alt="MergeEmptyVertical_after" width="50%" height="50%">

```WordUtil.bas```
------------------
```CreateWordDoc()```

Creates a Word document.

```ReplaceTagsWithContent()```

Fill content into a Word template by replacing tags ```<xxx>``` with content.

Example Cover Letter template and config are provided in ```config/``` folder.

Support for paragraph building with multiple sentences in same category (```P<xxx>``` tags in Excel config). This is done by extracting stock phrases from PhraseConfig sheet with a random starting phrase, and inserting the config data into the stock phrases.

```FillExcelForm.bas```
-----------------------
Fills highlighted cells in an Excel form, searching within a specified area for yellow colour (currently set at value of 13431551).

Before:

<img src="https://github.com/adrielyeung/macro-workbook/blob/main/img/FillExcelForm_Before.png" alt="FillExcelForm_Before" width="50%" height="50%">

After:

<img src="https://github.com/adrielyeung/macro-workbook/blob/main/img/FillExcelForm_After.png" alt="FillExcelForm_After" width="50%" height="50%">

```CopyColumnToNext.bas```
--------------------------
Copy the content of rightmost filled column to the next, increasing the header by 1 if it is a number/date.

Option to select:
1. Number of times to copy
2. If copy > 1 times, copy header only except last time (Useful for skipping through a few days, e.g. weekend)

Before:

<img src="https://github.com/adrielyeung/macro-workbook/blob/main/img/CopyColumnToNext_Before.png" alt="CopyColumnToNext_Before" width="50%" height="50%">

After:

- Case 1: Copy 1 time
  
  Prompt:
  
  Type 1 to copy 1 time
  
  <img src="https://github.com/adrielyeung/macro-workbook/blob/main/img/CopyColumnToNext_Case1.png" alt="CopyColumnToNext_Case1" width="50%" height="50%">
  
  After:
  
  <img src="https://github.com/adrielyeung/macro-workbook/blob/main/img/CopyColumnToNext_Case1_After.png" alt="CopyColumnToNext_Case1_After" width="50%" height="50%">
  
- Case 2: Copy 3 times, skipping except the last time (e.g. skip through the weekend)
  
  Prompt:
  
  Type 3 to copy 3 times
  
  <img src="https://github.com/adrielyeung/macro-workbook/blob/main/img/CopyColumnToNext_Case2_1.png" alt="CopyColumnToNext_Case2_1" width="50%" height="50%">
  
  Select "Yes" to set up the header (date) only
  
  <img src="https://github.com/adrielyeung/macro-workbook/blob/main/img/CopyColumnToNext_Case2_2.png" alt="CopyColumnToNext_Case2_2" width="50%" height="50%">
  
  After:
  
  <img src="https://github.com/adrielyeung/macro-workbook/blob/main/img/CopyColumnToNext_Case2_After.png" alt="CopyColumnToNext_Case2_After" width="50%" height="50%">

## Future developments
WordUtil.bas: Batch processing of config (maybe in CSV format).

FillExcelForm.bas: Create another sub for generating PDF copy and attaching to email.

