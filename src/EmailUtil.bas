Attribute VB_Name = "EmailUtil"
Sub GenPDFAndEmail()
'
' GenPDFAndEmail Sub
' Export the ActiveSheet of ActiveWorkbook as PDF,
' then create Outlook email with parameters, ready for send
'

    Dim ObjOutlook As Object, ObjEmail As Object
    Dim PdfName As String

    ' Export as PDF
    PdfName = PDFUtil.GenPDF("<Suffix>")

    ' Create Outlook object
    Set ObjOutlook = CreateObject("Outlook.Application")

    ' Create email object
    Set ObjEmail = ObjOutlook.CreateItem(olMailItem)

    ' Set parameters
    With ObjEmail
        .To = ""
        .Cc = ""
        .Subject = ""
        .Body = "Dear <ReceiverName>," & vbNewLine & vbNewLine & _
                "Attached please find my <document>."
                ' & ObjEmail.Body - to insert text signature directly
        .Attachments.Add (PdfName)
        .Display        ' Display the message in Outlook.
        ' Move to end of email to insert default signature manually
        SendKeys "^+{END}", True
        SendKeys "{END}", True
        SendKeys "{NUMLOCK}"
    End With

    ' Clear objects at end
    Set ObjEmail = Nothing
    Set ObjOutlook = Nothing
End Sub

Sub LeaveEmail_Dates()
'
' LeaveEmail_Dates Sub
' Create Outlook email to alert the team of leave between FromDate and ToDate
' Reading FromDate and ToDate from Excel config
'
'
    Dim FromDate As String, ToDate As String, FromAmPm As String, ToAmPm As String
    
    FromDate = Range("FromDate").Value
    ToDate = Range("ToDate").Value
    FromAmPm = Range("FromAmPm").Value
    ToAmPm = Range("ToAmPm").Value
    
    LeaveEmail FromDate, ToDate, FromAmPm, ToAmPm
End Sub

Sub LeaveEmail_LeaveLog()
'
' LeaveEmail_LeaveLog Sub
' Create Outlook email to alert the team of next available leave within a month
' Reading FromDate and ToDate from logging Excel of leave days,
' where each column represent a day and each row represent an employee
'
'
    Dim FromDate As String, ToDate As String, FromAmPm As String, ToAmPm As String, Name As String
    Dim LeaveLogWb As Workbook, NameRange As Range, DateRange As Range, LeaveRange As Range
    Dim RowNum As Long, ColNum As Long, i As Long, FromCol As Long, ToCol As Long
    
    ' Copy a backup in case of macro run failure
    FileCopy Range("LeaveLog").Value, Range("TempLog").Value
    
    Name = Range("Name").Value
    
    Set LeaveLogWb = Workbooks.Open(Range("TempLog").Value)
    Set NameRange = Range("Names")
    Set DateRange = Range("Dates")
    
    ' Find employee name record and current date column
    RowNum = NameRange.Find(Name, , xlValues, xlPart).Row
    ColNum = DateRange.Find(Date, , xlFormulas, xlPart).Column
    
    Set LeaveRange = Rows(RowNum)
    
    ' Loop today and next 31 days
    ' Half day: "A" for AM leave, "P" for PM leave
    ' Full day: "F" for annual leave, "CL" for comp leave and "BL" for birthday leave
    For i = ColNum To ColNum + 31
        ' Find from date
        If Len(Trim(Cells(RowNum, i))) > 0 Then
            FromDate = DateRange.Cells(1, i).Value
            If Cells(RowNum, i).Value = "A" Then
                FromAmPm = "AM"
                ToDate = FromDate
                ToAmPm = "AM"
                Exit For
            ElseIf Cells(RowNum, i).Value = "P" Then
                FromAmPm = "PM"
            End If
            
            ToCol = i
            i = i + 1
            
            ' Loop for whole leave period
            ' until next day without leave or weekend/holiday (greyed colour 12566463)
            While Cells(RowNum, i).Value = "F" Or Cells(RowNum, i).Value = "CL" Or _
                Cells(RowNum, i).Value = "BL" Or Cells(RowNum, i).Interior.Color = 12566463
                ' Only include day if it is a work day
                If Not Cells(RowNum, i).Interior.Color = 12566463 Then
                    ToCol = i
                End If
                i = i + 1
            Wend
            
            ' Handle any half day leaves at the end of period
            If Cells(RowNum, i).Value = "A" Then
                ToCol = i
                ToAmPm = "AM"
            End If
            
            ToDate = DateRange.Cells(1, ToCol)
            Exit For
        End If
    Next i
    
    ' Close the backup and do not save
    LeaveLogWb.Close False
    
    ' Can't find any leave if FromDate empty
    If Len(Trim(FromDate)) = 0 Then
        MsgBox "Cannot find leave for the next 31 days.", vbOKOnly, "No pending leave"
        Exit Sub
    End If
    
    ' Delete backup file
    ' Kill Range("TempLog").Value
    
    LeaveEmail FromDate, ToDate, FromAmPm, ToAmPm
End Sub

Private Sub LeaveEmail(FromDate As String, ToDate As String, FromAmPm As String, ToAmPm As String)
'
' LeaveEmail Sub
' Create Outlook email to alert leave between FromDate and ToDate
'
'
    Dim ObjOutlook As Object, ObjEmail As Object
    Dim Formatter As String, Name As String, EmailSubj As String, EmailTo As String, EmailBody As String
    
    Name = Range("Name").Value
    EmailTo = Range("EmailTo").Value
    EmailBody = Range("EmailBody").Value
    
    ' Take only first name (before first space)
    If InStr(1, Name, " ") > 0 Then
        Name = Left(Name, InStr(1, Name, " ") - 1)
    End If
    
    ' Create Outlook object
    Set ObjOutlook = CreateObject("Outlook.Application")
    
    ' Create email object
    Set ObjEmail = ObjOutlook.CreateItem(olMailItem)
    
    ' Check both dates are valid, and from date before/equal to date
    If Not IsDate(FromDate) Then
        MsgBox "Invalid From Date: " & FromDate, vbOKOnly, "Invalid From Date"
        Exit Sub
    End If
    If Not IsDate(ToDate) Then
        MsgBox "Invalid To Date: " & ToDate, vbOKOnly, "Invalid To Date"
        Exit Sub
    End If
    If CDate(FromDate) > CDate(ToDate) Then
        MsgBox "From Date: " & FromDate & vbNewLine & "after" & vbNewLine & _
            "To Date: " & ToDate, vbOKOnly, "From Date after To Date"
        Exit Sub
    ElseIf CDate(FromDate) = CDate(ToDate) And FromAmPm = "PM" And ToAmPm = "AM" Then
        MsgBox "From Date: " & FromDate & " " & FromAmPm & vbNewLine & "after" & vbNewLine & _
            "To Date: " & ToDate & " " & ToAmPm, vbOKOnly, "From Date after To Date"
        Exit Sub
    End If
    
    If Format(FromDate, "yyyy") = Format(ToDate, "yyyy") Then
        Formatter = "dd/mmm (ddd)"
    Else
        Formatter = "dd/mmm/yyyy (ddd)"
    End If
    
    FromDate = Format(FromDate, Formatter)
    ToDate = Format(ToDate, Formatter)
    
    If Len(Trim(FromAmPm)) > 0 Then
        FromAmPm = " " & FromAmPm
    End If
    
    If Len(Trim(ToAmPm)) > 0 Then
        ToAmPm = " " & ToAmPm
    End If
    
    If FromDate = ToDate Then
        EmailSubj = Name & " on leave " & FromDate & FromAmPm
    Else
        EmailSubj = Name & " on leave " & FromDate & FromAmPm & " to " & ToDate & ToAmPm
    End If
    
    ' Set parameters
    With ObjEmail
        .To = EmailTo
        .Subject = EmailSubj
        .Body = EmailBody & vbNewLine & vbNewLine
                ' & ObjEmail.Body - to insert text signature directly
        .Display        ' Display the message in Outlook.
        ' Move to end of email to insert default signature manually
        SendKeys "^+{END}", True
        SendKeys "{END}", True
        SendKeys "{NUMLOCK}"
    End With
    
    ' Clear objects at end
    Set ObjEmail = Nothing
    Set ObjOutlook = Nothing
End Sub

