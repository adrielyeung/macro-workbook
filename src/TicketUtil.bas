Attribute VB_Name = "TicketUtil"
Sub OpenCloseTicket()
' OpenCloseTicket Macro
' Open/Close ticket records searching for ad-hoc ticket number (format AHxxxxx, where x is a digit) as input.
'

'
    Dim TicketNum As String, TicketRow As Long
    
    Do While True
        TicketNum = InputBox("Enter ticket number (AHxxxxx):", "Ticket #")
        If IsNumeric(Mid(TicketNum, 3)) And Left(TicketNum, 2) = "AH" Then
            Exit Do
        End If
        If MsgBox("Please enter a valid ticket number (AHxxxxx).", vbOKCancel, "Number required") = vbCancel Then Exit Sub
    Loop
    
    TicketRow = ExcelUtil.MatchLast(TicketNum, Range("D:D"), 1)
    
    If Len(Range("I" & TicketRow).Value) = 0 Then
        Range("F" & TicketRow).Value = "Working in progress"
        Range("I" & TicketRow).Value = "Adriel"
        Range("J" & TicketRow).Value = "Other systems"
        Range("K" & TicketRow).Value = Date
        Range("L" & TicketRow).Value = Date + FindNextWorkday()
        Range("M" & TicketRow).Value = "50%"
        MsgBox "Ticket #" & TicketNum & " opened.", vbOKOnly, "Ticket opened"
    Else
        Range("F" & TicketRow).Value = "Closed"
        Range("G" & TicketRow).Value = Date
        Range("M" & TicketRow).Value = "100%"
        MsgBox "Ticket #" & TicketNum & " closed.", vbOKOnly, "Ticket closed"
    End If
End Sub
Sub RenewTicket()
'
' RenewTicket Macro
' Renew ad-hoc tickets records assigned to me that are due today or before.
'

'
    Dim CurRow As Integer, EndRow As Integer
    CurRow = ActiveSheet.UsedRange.Rows.Count
    EndRow = 2
    
    While CurRow >= EndRow
        ' Valid ticket
        If Len(Range("D" & CurRow).Value) > 0 Then
            ' Only check last 100 valid tickets
            If EndRow = 2 Then
                EndRow = CurRow - 100
            End If
            
            ' My ticket
            If Range("I" & CurRow).Value = "Adriel" Then
                ' Status not closed and due today or before
                If Range("L" & CurRow).Value <= Date + FindNextWorkday() - 1 _
                    And Not Range("F" & CurRow).Value = "Closed" Then
                    Range("L" & CurRow).Value = Date + FindNextWorkday()
                End If
            End If
        End If
        CurRow = CurRow - 1
    Wend
    
    MsgBox "Renewed your tickets successfully.", vbOKOnly, "Tickets renewed"
    
End Sub

Function FindNextWorkday()
'
' FindNextWorkday Function
' Find next work day after tomorrow.
'

'
    If Weekday(Date, vbMonday) >= 4 Then
        FindNextWorkday = 4
    Else
        FindNextWorkday = 2
    End If
End Function

