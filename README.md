# macro-workbook
A collection of VBA Excel macros that I use for utility.

## How to use
Suggest to import all the ```.bas``` files into your Personal Macro Workbook for use across different Excel files.

## Description
### Cover Letter Template Generation 1 - Write Cover Letter
The idea is to single/batch generate cover letters based on a template, loading the information from a job application Excel (```config/JobApplication.xlsx```) and a experience list Excel (```config/ProjectList.xlsx```).

1. Single generation (```WordUtil.Single_ReplaceTagsWithContent```)
2. Batch generation (```WordUtil.Batch_ReplaceTagsWithContent```): copy job title, company name and required skills of each record with Status = "Prepare" from ```config/JobApplication.xlsx``` into ```config/CoverLetterGenerator.xlsx```. Successful records will have Status changed to "GenPDF" ready for [part 2](#cover-letter-template-generation-2---export-to-pdf), Word file name logged into "Message" column of ```config/JobApplication.xlsx```. Error messages will be logged into "Message" column too.

Both call the function ```WordUtil.ReplaceTagsWithContent``` which fills a copy of ```config/CoverLetterTemplate.docx``` by replacing tags ```<xxx>``` with content from ```config/CoverLetterGenerator.xlsx```, relevant records of ```config/ProjectList.xlsx```).

Support for:
1. Paragraph building with multiple phrases in same category (```P<xxx>``` tags in Excel config). This is done by extracting stock phrases from PhraseConfig sheet with a random starting phrase, and inserting the config data into the stock phrases.
2. Paragraph building drawing relevant experiences for each skill (```L<xxx>``` tags in Excel config). This is done by fetching records in ```config/ProjectList.xlsx``` searching in the "Skill" column.

See below for the flow of the program:

<img src="https://github.com/adrielyeung/macro-workbook/blob/main/img/CoverLetterGenerator_Flow.png" alt="Cover Letter Generator Flow" width="50%" height="50%">

### Cover Letter Template Generation 2 - Export to PDF
After reviewing the macro-generated letter in part 1, batch generate PDF by calling ```PDFUtil.Batch_ExportWordAsPDF```. This looks for all records with status = "GenPDF" and export those Word documents into PDF. After success, will have Status changed to "Done", else error message will be logged into "Generation Message" column.

Steps:

1. Template

<img src="https://github.com/adrielyeung/macro-workbook/blob/main/img/CoverLetterGenerator_Step1.png" alt="Cover Letter Generator Step 1" width="50%" height="50%">

2. Word document

<img src="https://github.com/adrielyeung/macro-workbook/blob/main/img/CoverLetterGenerator_Step2.png" alt="Cover Letter Generator Step 2" width="50%" height="50%">

3. PDF document

<img src="https://github.com/adrielyeung/macro-workbook/blob/main/img/CoverLetterGenerator_Step3.png" alt="Cover Letter Generator Step 3" width="50%" height="50%">

```FindKeyword.bas```
------------------------
This is a macro used to search for a number of keywords in the config file (with comma-separated keyword values).

For each row, the found keyword(s) in search range will be put into the first empty column (Result Column), as a comma-separated list.

A relevancy score is calculated and placed into the second empty column. The results are sorted according to descending relevancy.

Inputs (via InputBox / MsgBox):
1. Config file path
2. Search Range (e.g. "A1:B2" Range in Excel)
3. Whether to use equal scoring for each keyword, or weighted scoring (weights will be added in descending order of each word in the config file). The final score would be used to sort the results.

All 3 inputs are compulsory.

```ExcelUtil.bas```
-------------------
This file contains macros which operate in Excel files.

### 1. FillExcelForm
Fills highlighted cells in an Excel form, searching within a specified area for yellow colour (currently set at value of 13431551).

Before:

<img src="https://github.com/adrielyeung/macro-workbook/blob/main/img/FillExcelForm_Before.png" alt="FillExcelForm_Before" width="50%" height="50%">

After:

<img src="https://github.com/adrielyeung/macro-workbook/blob/main/img/FillExcelForm_After.png" alt="FillExcelForm_After" width="50%" height="50%">

### 2. CopyDataToSheet
Copy the data from "From" named area to "To" named area and set up Status in "PrepareStatus" named area as "Prepare" for newly copied rows.

### 3. CopyColumnToNext
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

### 4. MergeEmptyVertical
This is a macro used to merge empty cells in each column in a selected range with the nearest non-empty cell above.

Additionally, can output the number of rows merged for each non-empty cell to a specified column (requires the non-empty cell to be on same row, as in column "D" in below screenshot).

Before:

<img src="https://github.com/adrielyeung/macro-workbook/blob/main/img/MergeEmptyVertical_before.png" alt="MergeEmptyVertical_before" width="50%" height="50%">

After:

<img src="https://github.com/adrielyeung/macro-workbook/blob/main/img/MergeEmptyVertical_after.png" alt="MergeEmptyVertical_after" width="50%" height="50%">

```WordUtil.bas```
------------------
This file contains Excel macros whose output is in Word files.

### 1. CreateWordDoc
Creates a Word document.

### 2. Batch_, Single_ and ReplaceTagsWithContent Function
Part of the [Cover Letter Generator project](#cover-letter-template-generation-1---write-cover-letter).

```PDFUtil.bas```
-----------------
This file contains Excel macros whose output is in PDF files.

### 1. Batch_, Single_ and ExportWordAsPDF
Part of the [Cover Letter Generator project](#cover-letter-template-generation-1---write-cover-letter).

## Future developments
Cover Letter Generation: Allow selection of number of items to copy from project template.

FillExcelForm.bas: Create another sub for generating PDF copy and attaching to email.
