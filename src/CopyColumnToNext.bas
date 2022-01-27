Attribute VB_Name = "CopyColumnToNext"
Sub CopyColumnToNext()
'
' CopyColumnToNext Macro
' Copy the content of rightmost filled column to the next,
' increasing the header by 1 if it is a number/date.
'
' Option to select:
' 1) Number of times to copy
' 2) If copy > 1 times, copy header only except last time
' (Useful for skipping through a few days, e.g. weekend)
'
' Keyboard Shortcut: Ctrl+k
'
    Dim LastColInd As Integer, CopyTimes As Variant, HeaderVal As Variant
    Dim CopyHeaderOnly As String
    
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
        
    HeaderVal = Cells(1, LastColInd).Value
    
    For i = 1 To CInt(CopyTimes)
        ' Start a new column
        LastColInd = LastColInd + 1
        
        ' Increment header by 1
        Cells(1, LastColInd).Value = Cells(1, LastColInd - 1).Value + 1
        
        ' Copy remaining rows
        If CopyHeaderOnly = vbNo Or i = CInt(CopyTimes) Then
            Range(Cells(2, LastColInd), Cells(ActiveSheet.UsedRange.Rows.Count, LastColInd)).Value = _
                Range(Cells(2, LastColInd - i), Cells(ActiveSheet.UsedRange.Rows.Count, LastColInd - i)).Value
        End If
        
        ' Autofit filled column
        ActiveSheet.Columns(LastColInd).AutoFit
    Next i
    
    ' Save workbook
    ActiveWorkbook.Save
End Sub

