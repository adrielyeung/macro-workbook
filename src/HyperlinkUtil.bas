Attribute VB_Name = "HyperlinkUtil"
Sub AddHyperlinkToColumn()
'
' AddHyperlink Macro
' Batch add hyperlinks to each row in a column (skipping first row).
' Link to another column in the same row.
'

'
    Dim i As Integer
    
    For i = 2 To ActiveSheet.UsedRange.Rows.Count
    
        Range("B" & i).Value = "SMS"
        
        ActiveSheet.Hyperlinks.Add Anchor:=Range("B" & i), Address:="", SubAddress:= _
            "Content!D" & i, TextToDisplay:="SMS"
        
        Range("B" & i).Font.Name = "Arial"
    Next i
    
End Sub
Sub NavigationHyperlink()
'
' NavigationHyperlink Macro
' Add a new Sheet for each record, and
' add hyperlinks from specific column to/from cell "A1" of new Sheet to "content page" Sheet (Content).
'

'
    Dim i As Integer
    
    For i = 2 To ActiveSheet.UsedRange.Rows.Count
        Range("C" & i).Value = "Link-" & i - 1
        
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = Sheets("Content").Range("C" & i).Value
        
        ActiveSheet.Hyperlinks.Add Anchor:=Range("A1"), Address:="", SubAddress:= _
            "Content!C" & i, TextToDisplay:="Back to main"
            
        Range("A1").Font.Name = "Arial"
        
        Sheets("Content").Select
        
        Sheets("Content").Hyperlinks.Add Anchor:=Range("C" & i), Address:="", SubAddress:= _
            "'" & Range("C" & i).Value & "'!A1", TextToDisplay:=Range("C" & i).Value
            
        Range("C" & i).Font.Name = "Arial"
    Next i
End Sub

Sub UpdateHyperlink()
'
' UpdateHyperlink Macro
' Update each navigation hyperlink to "content page" sheet (Content).

'
    Dim i As Integer
    
    For i = 2 To Sheets.Count
        Sheets(i).Select
        
        Sheets(i).Hyperlinks.Add Anchor:=Sheets(i).Range("A1"), Address:="", SubAddress:= _
            "Content!B" & Left(Sheets(i).Name, 2) + 1, TextToDisplay:="Back to main"
        
        Range("A1").Font.Name = "Arial"
    Next i
    
    Sheets(1).Select
End Sub
