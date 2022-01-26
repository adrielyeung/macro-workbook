Attribute VB_Name = "FillExcelForm"

Sub FillExcelForm()
'
' FillExcelForm Macro
' Fill in Excel form Document based on highlighted cells.
'
' Keyboard Shortcut: Ctrl+j
'
    ' Fill in basic info
    Range("B14:D15").Value = ""
    Range("F14:H15").Value = ""
    Range("B16:D17").Value = ""
    
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
