Attribute VB_Name = "TicketUtil"
Sub OpenCloseTicket()
' OpenCloseTicket Macro
' Open/Close ticket records searching for ticket number as input.
'

'
    Dim TicketNum As String, TicketRow As Long
    
    Do While True
        TicketNum = InputBox("Enter ticket number:", "Ticket #")
        If IsNumeric(TicketNum) Then
            Exit Do
        End If
        If MsgBox("Please enter a valid ticket number.", vbOKCancel, "Number required") = vbCancel Then Exit Sub
    Loop
    
    TicketRow = ExcelUtil.MatchLast(TicketNum, Range("TicketNum"), 1)
    
    ' Ticket not found, exit
    If TicketRow <= 0 Then
        MsgBox "Ticket #" & TicketNum & " not found.", vbOKOnly, "Ticket not found"
        Exit Sub
    End If
    
    ' Check if ticket is already opened (assigned to person)
    If Len(Range("Assignee").Cells(TicketRow, 1).Value) = 0 Then
        Range("Status").Cells(TicketRow, 1).Value = "Working in progress"
        Range("Assignee").Cells(TicketRow, 1).Value = Range("Name").Value
        Range("UpdateDateTime").Cells(TicketRow, 1).Value = Date
        Range("ExpectedDateTime").Cells(TicketRow, 1).Value = Date + FindNextWorkday()
        MsgBox "Ticket #" & TicketNum & " opened.", vbOKOnly, "Ticket opened"
    Else
        Range("Status").Cells(TicketRow, 1).Value = "Closed"
        Range("UpdateDateTime").Cells(TicketRow, 1).Value = Date
        MsgBox "Ticket #" & TicketNum & " closed.", vbOKOnly, "Ticket closed"
    End If
End Sub
Sub RenewTicket()
'
' RenewTicket Macro
' Renew tickets records assigned to me that are due today or before.
'

'
    Dim CurRow As Integer, EndRow As Integer
    CurRow = ActiveSheet.UsedRange.Rows.count
    EndRow = 2
    
    While CurRow >= EndRow
        ' Valid ticket
        If Len(Range("TicketNum").Cells(CurRow, 1).Value) > 0 Then
            ' Only check last 100 valid tickets
            If EndRow = 2 And CurRow - 100 > 2 Then
                EndRow = CurRow - 100
            End If
            
            ' My ticket
            If Range("Assignee").Cells(CurRow, 1).Value = Range("Name").Value Then
                ' Status not closed and due today or before
                If Range("ExpectedDateTime").Cells(CurRow, 1).Value <= Date + FindNextWorkday() - 1 _
                    And Not Range("Status").Cells(CurRow, 1).Value = "Closed" Then
                    Range("ExpectedDateTime").Cells(CurRow, 1).Value = Date + FindNextWorkday()
                End If
            End If
        End If
        CurRow = CurRow - 1
    Wend
    
    MsgBox "Renewed your tickets successfully.", vbOKOnly, "Tickets renewed"
    
End Sub

Sub ListEmailSubject()
'
' ListEmailSubject Sub
' List out all email subjects from last run time to now containing certain text
'

    Dim outApp As Outlook.Application
    Dim outNameS As Outlook.Namespace
    Dim outAttachment As Outlook.Attachment
    Dim outFolderToCheck As Outlook.Folder
    Dim outItem As Object

    Dim loginEmail As String, fromEmails As Range, containKeywords As Range, filterOutKeywords As Range, flags As Range, flagKeywords As Range
    Dim mainFolder As String, flag As Range, flagKeyword As Range
    Dim emailSubject As String, emailBody As String, emailSender As String, filterOutKeyword As Range
    Dim outMessage As String
    Dim lastRunDateTime As Date
    Dim keyword As Range, sender As Range
    Dim sh As Worksheet, emailSh As Worksheet
    Dim searchLimit As Long, CurRow As Long, firstRow As Long, count As Long
    Dim validSender As Boolean, allFilterKeywordRemoved As Boolean
    
    Set sh = ActiveWorkbook.Worksheets("Config")
    Set emailSh = ActiveWorkbook.Worksheets("Email")

    Set outApp = Outlook.Application
    Set outNameS = outApp.GetNamespace("MAPI")

    'Get email account and folder parameters from Email tab
    loginEmail = sh.Range("LoginEmail")
    Set fromEmails = sh.Range("FromEmail")
    mainFolder = sh.Range("Folder")
    Set containKeywords = sh.Range("Contains")
    Set filterOutKeywords = sh.Range("FilterOut")
    searchLimit = sh.Range("SearchLimit")
    Set flags = sh.Range("Flag")
    Set flagKeywords = sh.Range("FlagKeywords")
    lastRunDateTime = CDate(sh.Range("LastRunDateTime"))
    CurRow = emailSh.UsedRange.Rows.count
    firstRow = emailSh.Range("TicketNum").Cells(CurRow, 1).Value + 1
    count = 0

    'Assign the folder to a variable
    Set outFolderToCheck = outNameS.Folders(loginEmail).Folders(mainFolder)

    'Check if there is any email in the folder
    If outFolderToCheck.Items.count > 0 Then
        
        'Loop through in each item in the folder (including mails, meetings, tasks)
        For Each outItem In outFolderToCheck.Items
            
            'Check if item is a mail
            If TypeOf outItem Is MailItem Then
            
                ' Only process messages after last run
                If outItem.ReceivedTime > lastRunDateTime Then
                    
                    emailSubject = outItem.Subject
                    
                    ' Remove all FilterOut keywords from subject
                    Do
                        allFilterKeywordRemoved = True
                        For Each filterOutKeyword In filterOutKeywords
                            If InStr(1, Left(emailSubject, Len(filterOutKeyword.Value)), filterOutKeyword.Value, vbTextCompare) Then
                                emailSubject = Mid(emailSubject, Len(filterOutKeyword.Value) + 2)
                                allFilterKeywordRemoved = False
                                Exit For
                            End If
                        Next filterOutKeyword
                    Loop Until allFilterKeywordRemoved
                    
                    ' Check if from valid sender
                    emailSender = outItem.sender
                    validSender = False
                    For Each sender In fromEmails
                        If InStr(1, emailSender, sender, vbTextCompare) > 0 Then
                            validSender = True
                            Exit For
                        End If
                    Next sender
                    
                    If validSender Then
                        emailBody = outItem.Body
                        
                        ' Check if email body contains keyword in "Contains" section in the first "SearchLimit" chars
                        For Each keyword In containKeywords
                            ' Also check that subject do not already exist
                            If Len(keyword.Value) > 0 And InStr(1, Left(emailBody, searchLimit), keyword.Value, vbTextCompare) > 0 And _
                                MatchLast(emailSubject, emailSh.Range("Subject"), CurRow, 1) = 0 Then
                                CurRow = CurRow + 1
                                count = count + 1
                                emailSh.Range("TicketNum").Cells(CurRow, 1).Value = _
                                    emailSh.Range("TicketNum").Cells(CurRow - 1, 1).Value + 1
                                emailSh.Range("Subject").Cells(CurRow, 1).Value = emailSubject
                                emailSh.Range("ReceivedDateTime").Cells(CurRow, 1).Value = outItem.ReceivedTime
                                
                                For Each flag In flags
                                    Set flagKeyword = flag.Offset(0, 1)
                                    While Len(flagKeyword.Value) > 0
                                        If InStr(1, emailSubject, flagKeyword.Value, vbTextCompare) > 0 Then
'                                            InStr(1, Left(emailBody, searchLimit), flagKeyword.Value, vbTextCompare) > 0 Then
                                            emailSh.Range(flag.Value).Cells(CurRow, 1) = "Y"
                                        End If
                                        Set flagKeyword = flagKeyword.Offset(0, 1)
                                    Wend
                                Next flag
                            End If
                        Next keyword
                    End If
                End If
            End If
        Next outItem

    End If
    
    ' Update last run date time to now
    sh.Range("LastRunDateTime") = Now
    
    outMessage = "Complete ticket logging from " & Format(lastRunDateTime, "yyyy-mm-dd hh:mm:ss") & " to " & _
        Format(Now, "yyyy-mm-dd hh:mm:ss") & vbNewLine & vbNewLine & _
        "New tickets count: " & count & vbNewLine & vbNewLine
    
    If count > 0 Then
        outMessage = outMessage & "#" & firstRow & " to #" & emailSh.Range("TicketNum").Cells(CurRow, 1).Value
    End If
    
    MsgBox outMessage
    
End Sub

Function MatchLast(Lookupvalue As String, LookupRange As Range, StartRowNumber As Long, ColumnNumber As Integer) As Long
' MatchLast Function
' Returns last cell row number in LookupRange, ColumnNumber'th column containing Lookupvalue
'
    Dim i As Long
    For i = StartRowNumber To 1 Step -1
        If Lookupvalue = LookupRange.Cells(i, 1) Then
            MatchLast = i
            Exit Function
        End If
    Next i
End Function

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

