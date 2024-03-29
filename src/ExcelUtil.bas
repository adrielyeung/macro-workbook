Attribute VB_Name = "ExcelUtil"
Sub FillExcelForm()
'
' FillExcelForm Macro
' Fill in Excel form Document based on highlighted cells.
'
' Keyboard Shortcut: Ctrl+j
'
    ' Fill in basic info
'    Range("B14:D15").Value = ""
'    Range("F14:H15").Value = ""
'    Range("B16:D17").Value = ""
    
    Dim CalCell As Range
    
    ' Search for coloured cell (to be filled) in the calendar area
    ' Look for values allowed
    ' Default select first value for each validation dropdown
    For Each CalCell In Range("B24:H59")
        If CalCell.DisplayFormat.Interior.Color = 13431551 Then
            CalCell.Value = Range(CalCell.Validation.Formula1).Cells(1, 1).Value
        End If
    Next CalCell

    ActiveWorkbook.Save
    
    MsgBox "Filled in successfully with default values, please check.", vbOKOnly, "Success"
End Sub

Sub CopyDataToSheet()
'
' CopyDataToSheet Macro
' Copy the data from "From" named area to "To" named area
' And set up Status as "Prepare" for newly copied rows
'
'
    Dim CopyTo As Range, CopyFrom As Range, PrepareStatus As Range
    Dim i As Integer
    Set CopyTo = ActiveWorkbook.Names("CopyTo").RefersToRange
    Set CopyFrom = ActiveWorkbook.Names("CopyFrom").RefersToRange
    Set PrepareStatus = ActiveWorkbook.Names("PrepareStatus").RefersToRange
    
    ' Copy from From area to To area
    CopyFrom.Copy
    CopyTo.Worksheet.Activate
    CopyTo.Range(Cells(1, 1), Cells(CopyFrom.Rows.Count, CopyFrom.Columns.Count)).PasteSpecial Paste:=xlPasteAll
    
    CopyTo.Rows.AutoFit
    
    ' Copy formatting
    PrepareStatus.Cells(2, 1).Copy
    
    ' Set end cell for "Status" and "Message" cols
    With PrepareStatus.Cells(CopyFrom.Rows.Count, 1)
        .Value = "End"
        .PasteSpecial Paste:=xlPasteFormats
        .Offset(0, 1).Value = "End"
        .Offset(0, 2).Value = "End"
    End With
    
    ' Set newly added rows with status "Prepare" until found previously done rows
    For i = CopyFrom.Rows.Count - 1 To 2 Step -1
        If IsEmpty(PrepareStatus.Cells(i, 1)) Or PrepareStatus.Cells(i, 1).Value = "End" Then
            With PrepareStatus.Cells(i, 1)
                .Value = "Prepare"
                .PasteSpecial Paste:=xlPasteFormats
                .Offset(0, 1).Value = ""
                .Offset(0, 2).Value = ""
            End With
        Else
            Exit For
        End If
    Next i
    
    ' Filter status = "Prepare" rows for work
    PrepareStatus.AutoFilter Field:=PrepareStatus.Column, Criteria1:="Prepare", Operator:=xlOr, Criteria2:="End"
    ActiveWorkbook.Save
End Sub

Sub CopyColumnToNext()
'
' CopyColumnToNext Macro
' Copy the content of rightmost filled non-coloured column to the next,
' increasing the header by 1 if it is a number/date.
'
' Option to select:
' 1) Number of times to copy
' 2) If copy > 1 times, copy header only except last time
' (Useful for skipping through a few days, e.g. weekend)
'
' Keyboard Shortcut: Ctrl+k
'
    Dim LastColInd As Integer, CopyTimes As Variant
    Dim CopyHeaderOnly As String
    Dim i As Long, j As Long
    
    LastColInd = ActiveSheet.UsedRange.Columns.Count
    
    Do While True
        CopyTimes = InputBox("Please enter number of times you want to copy the last column for:", _
            "Copy Times", "1")
        If IsNumeric(CopyTimes) Then
            Exit Do
        End If
        If MsgBox("Please enter a number.", vbOKCancel, "Number required") = vbCancel Then Exit Sub
    Loop
    
    If CInt(CopyTimes) > 1 Then
        CopyHeaderOnly = MsgBox("Copy header only except last time?", vbQuestion + vbYesNo, "Copy header only")
    Else
        CopyHeaderOnly = vbYes
    End If
    
    For j = 1 To ActiveSheet.UsedRange.Rows.Count
        ' Check for cells with white colour
        If Cells(j, LastColInd).EntireRow.Hidden = False And Cells(j, LastColInd).Interior.Color = 16777215 Then
            While Len(Trim(Cells(j, LastColInd).Value)) = 0 And LastColInd > 0
                LastColInd = LastColInd - 1
            Wend
            
            Exit For
        End If
    Next j
    
    For i = 1 To CInt(CopyTimes)
        ' Start a new column
        LastColInd = LastColInd + 1
        
        ' Increment header by 1 if not written
        If Len(Trim(Cells(1, LastColInd).Value)) = 0 Then
            Cells(1, LastColInd).Value = Cells(1, LastColInd - 1).Value + 1
        End If
        
        ' Copy remaining rows
        If CopyHeaderOnly = vbNo Or i = CInt(CopyTimes) Then
            For j = 2 To ActiveSheet.UsedRange.Rows.Count
                If Cells(j, LastColInd).EntireRow.Hidden = False And Cells(j, LastColInd).Interior.Color = 16777215 Then
                    Cells(j, LastColInd).Value = Cells(j, LastColInd - i).Value
                End If
            Next j
        End If
        
        ActiveSheet.Columns(LastColInd).AutoFit
    Next i
    
    ActiveWorkbook.Save
End Sub

Sub MergeEmptyVertical()
'
' MergeEmptyVertical Macro
' Merge empty cells vertically for each column (but do not center).
' For each merge, count number of cells merged and output to designated column.
'

'
    Dim ColInd As Long, RowInd As Long, ColStart As Long, ColEnd As Long, RowStart As Long, RowEnd As Long, MergeCount As Long
    Dim StartCell As String, EndCell As String, OutputColumn As String
    Dim FirstMerge As Boolean
    
    ' Output column for number of rows merged
    OutputColumn = Application.InputBox(Title:="Output column", Prompt:="Please provide merged column count output column (e.g. A)?", Type:=2, Default:="H")
    
    ColStart = Selection.Columns(1).Column
    ColEnd = Selection.Columns.Count + ColStart - 1
    RowStart = Selection.Rows(1).Row
    RowEnd = Selection.Rows.Count + RowStart - 1
    
    ' For each column in used range
    For ColInd = ColStart To ColEnd
        ' First time merge, no previous merge
        FirstMerge = True
        MergeCount = 0
        StartCell = Cells(RowStart, ColInd).Address()
        For RowInd = RowStart To RowEnd
            ' Find next non-empty cell
            If Not IsEmpty(Cells(RowInd, ColInd)) Then
                ' Finish previous merge
                If Not FirstMerge Then
                    EndCell = Cells(RowInd - 1, ColInd).Address()
                    ' Print merged column from start to end
                    If Len(OutputColumn) = 1 Then
                        Cells(Range(StartCell).Row, Range(OutputColumn & 1).Column) = MergeCount
                    End If
                    Range(StartCell, EndCell).Merge
                    MergeCount = 0
                End If
                FirstMerge = False
                StartCell = Cells(RowInd, ColInd).Address()
            End If
            MergeCount = MergeCount + 1
        Next
        ' Return to last row index
        RowInd = RowInd - 1
        If Len(OutputColumn) = 1 Then
            Cells(Range(StartCell).Row, Range(OutputColumn & 1).Column) = MergeCount
        End If
        EndCell = Cells(RowEnd, ColInd).Address()
        Range(StartCell, EndCell).Merge
        MergeCount = 0
    Next
End Sub

Function RangeExists(R As String) As Boolean
' RangeExists Function
' Test if Range R exists
'
    Dim Test As Range
    On Error Resume Next
    Set Test = ActiveSheet.Range(R)
    RangeExists = (Err.Number = 0)
End Function

Function SheetExists(S As String, Optional Wb As Workbook) As Boolean
' SheetExists Function
' Test if Sheet S in Wb if given, or ActiveWorkbook if Wb not given
'
    Dim Test As Worksheet
    If Wb Is Nothing Then
        Wb = ActiveWorkbook
    End If
    On Error Resume Next
    Set Test = Wb.Sheets(S)
    SheetExists = (Not Test Is Nothing)
End Function

Function MatchLast(Lookupvalue As String, LookupRange As Range, ColumnNumber As Integer) As Long
' MatchLast Function
' Returns last cell row number in LookupRange, ColumnNumber'th column containing Lookupvalue
'
    Dim i As Long
    For i = LookupRange.Columns(ColumnNumber).Cells.Count To 1 Step -1
        If Lookupvalue = LookupRange.Cells(i, 1) Then
            MatchLast = i
            Exit Function
        End If
    Next i
End Function


