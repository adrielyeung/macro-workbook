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

