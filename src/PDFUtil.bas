Attribute VB_Name = "PDFUtil"
Sub Single_ExportWordAsPDF()
'
' Single_ExportWordAsPDF Sub
' Single call of ExportWordAsPDF function to generate batch of PDFs from Word docs.
'

    Dim ExportPath As String, OutMsg As String
    
    If ExcelUtil.RangeExists("ExportPath") Then
        ExportPath = Range("ExportPath").Value
    Else
        Do While True
            ExportPath = InputBox("Please enter the full path of the Word document.", "Export Path")
            If Dir(ExportPath) <> "" Then
                Exit Do
            Else
                If MsgBox("Please enter a valid path.", vbOKCancel, "Path invalid") = vbCancel Then Exit Sub
            End If
        Loop
    End If
    
    OutMsg = ExportWordAsPDF(ExportPath)
    
    If Left(OutMsg, 1) = "!" Then
        OutMsg = Mid(OutMsg, 1)
    Else
        OutMsg = "PDF Generation SUCCESS, at: " & OutMsg
    End If
    
    MsgBox Prompt:=OutMsg, Title:="Finish Generation"

End Sub

Sub Batch_ExportWordAsPDF()
'
' Batch_ExportWordAsPDF Sub
' Batch call of ExportWordAsPDF function to generate batch of PDFs from Word docs, finding for records in path FeedPath
' with PrepareStatus = "GenPDF"
'

    Dim Status As Range, PrepareStatus As Range
    Dim i As Integer, SucCount As Integer, FailCount As Integer
    Dim OutMsg As String

    ' PrepareStatus are the cells holding the processing status
    Set PrepareStatus = ActiveWorkbook.Names("PrepareStatus").RefersToRange
    
    ' Error msg from each individual run
    OutMsg = ""
    SucCount = 0
    FailCount = 0
    
    For Each Status In PrepareStatus
        ' Extract all "GenPDF" status
        If Status.Value = "GenPDF" Then
            OutMsg = ExportWordAsPDF(Status.Offset(0, 1).Value)
            If Left(OutMsg, 1) <> "!" Then
                Status.Value = "GenPDF"
                SucCount = SucCount + 1
            Else
                FailCount = FailCount + 1
            End If
            Status.Offset(0, 2).Value = OutMsg
        ' Hit the end of file, finish
        ElseIf Status.Value = "End" Then
            Exit For
        End If
    Next Status

Output:
    If FailCount > 0 Then
        MsgBox Prompt:="Finish generating all PDFs WITH ERROR, please check." & vbNewLine _
        & "Success count: " & SucCount & vbNewLine _
        & "Failed count: " & FailCount, Title:="ERROR in PDF Generation"
    ElseIf SucCount = 0 Then
        MsgBox Prompt:="No PDFs generated.", Title:="No PDF Generation"
    Else
        MsgBox Prompt:="Finish generating all PDFs SUCCESS." & vbNewLine _
        & "Count: " & SucCount, Title:="Finish PDF Generation"
    End If

End Sub

Private Function ExportWordAsPDF(Path As String) As String
'
' ExportWordAsPDF Function
' Export Word doc specified by Path as PDF
'
'
    Dim Direc As String, Filename As String, ExtIndex As Integer
    Dim WordApp As Word.Application
    Set WordApp = New Word.Application
    
    ExportWordAsPDF = "!PDF Generation with ERROR. Please check code."
    
    Direc = Left(Path, InStrRev(Path, Application.PathSeparator))
        
    ExtIndex = InStrRev(Path, ".")
    If ExtIndex > 0 Then
        Filename = Mid(Path, Len(Direc) + 1, ExtIndex - Len(Direc) - 1)
    Else
        Filename = Mid(Path, Len(Direc) + 1)
    End If
    
    ChDir Direc
    
    WordApp.Documents.Open Path
    WordApp.Visible = False
    
    WordApp.ActiveDocument.ExportAsFixedFormat OutputFileName:= _
        Direc & Filename & ".pdf" _
        , ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
        wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=1, To:=1, _
        Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
        CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
        BitmapMissingFonts:=True, UseISO19005_1:=False
    
    ExportWordAsPDF = Direc & Filename & ".pdf"
End Function

