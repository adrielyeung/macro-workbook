Attribute VB_Name = "WordUtil"
Sub CreateWordDoc()
'
' CreateWordDoc Sub
' Creates a Word document
'

'
    Dim WordApp As Word.Application
    Set WordApp = New Word.Application
    
    With WordApp
        .Visible = True
        .Activate
        .Documents.Add
    End With

End Sub

Sub ReplaceTagsWithContent()
'
' ReplaceTagsWithContent Sub
' Replace placeholder tags <xxx> with actual content according to Excel config
'

'
    Dim Tags As Range, Tag As Range, Content As Range
    Dim Missing As Boolean
    Dim Template As String, NewFile As String, Direc As String
    Dim CompName As String
    Dim Prefix As String, Suffix As String, Category As String, LastCategory As String
    Dim Phrase As String, Paragraph As String, TagName As String, TagContent As String
    Dim Random As Double, PhraseRow As Long, FirstPhraseInd As Long, LastPhraseInd As Long, TagRow As Integer
    Dim WordApp As Word.Application
    Set WordApp = New Word.Application
    
    Set Tags = Range("Tags")
    Missing = False
    
    ' Checking if any content field is missing
    For Each Tag In Tags
        
        If Not IsEmpty(Tag.Value) Then
            Set Content = Tag.Offset(0, 1)
            
            If Left(Tag.Value, 1) <> "P" And IsEmpty(Content.Value) Then
                ' Highlight missing cell
                Content.Interior.ColorIndex = 6
                ' Select first missing cell
                If Not Missing Then
                    Content.Select
                End If
                Missing = True
            Else
                Content.Interior.ColorIndex = 2
            End If
        End If
        
    Next Tag
    
    If Missing Then
        MsgBox Prompt:="Please fill in highlighted fields.", Buttons:=vbOKOnly, Title:="Missing Fields"
        Exit Sub
    End If
    
    ' Set up template file
    If RangeExists("Template") Then
        Template = Range("Template").Value
    Else
        Template = InputBox("Please enter template file path", "Template", "C:\Users\Adriel\OneDrive\Documents\Careers\CoverLetter\CoverLetterTemplate.docx")
    End If
    Direc = Left(Template, InStrRev(Template, Application.PathSeparator))
    
    ' Duplicate template file for filling in
    If RangeExists("FileNamePrefix") Then
        Prefix = Range("FileNamePrefix")
    Else
        Prefix = "Document_"
    End If
    
    If RangeExists("CompName") Then
        NewFile = Prefix & Range("CompName").Value & "_" & Format(Now(), "yyyymmdd_hhmmss") & ".docx"
    Else
        NewFile = Prefix & Format(Now(), "yyyymmdd_hhmmss") & ".docx"
    End If
    
    FileCopy Template, Direc & NewFile
    
    ChDir Direc
    WordApp.Documents.Open Direc & NewFile
    
    WordApp.Visible = True
    
    ' Find and replace tags in Word template
    If SheetExists("DateConfig", ActiveWorkbook) Then
        Suffix = Application.WorksheetFunction.VLookup(CInt(Format(Date, "d")), ActiveWorkbook.Worksheets("DateConfig").Range("UsedDateConfig"), 2, True)
    End If
    
    FindAndReplace WordApp.ActiveDocument, Range("Date").Value, Format(Date, "d") & Suffix & Format(Date, " mmmm, yyyy.")
    
    Randomize
    Random = Rnd()
    Paragraph = ""
    For Each Tag In Tags
        
        Set Content = Tag.Offset(0, 1)
        
        If Not IsEmpty(Tag.Value) Then
            ' If tag starts with "P", loop phrases to load stock phrases
            If Left(Tag.Value, 1) = "P" Then
                If Not IsEmpty(Content.Value) Then
                    Category = Right(Tag.Value, Len(Tag.Value) - 1)
                    ' Changed category, add last category into template
                    If Category <> LastCategory And LastCategory <> "" Then
                        Randomize
                        Random = Rnd()
                        ' Remove the last space of Paragraph before inserting into template
                        FindAndReplace WordApp.ActiveDocument, LastCategory, Left(Paragraph, Len(Paragraph) - 1)
                        Paragraph = ""
                    End If
                    
                    ' Calculate new random index of stock phrases to start from
                    If Category <> LastCategory Then
                        FirstPhraseInd = Application.WorksheetFunction.Match(Category, ActiveWorkbook.Worksheets("PhraseConfig").Range("PhraseTags"), 0)
                        LastPhraseInd = MatchLast(Category, ActiveWorkbook.Worksheets("PhraseConfig").Range("PhraseTags"), 1)
                        
                        PhraseRow = Application.WorksheetFunction.RoundDown(Random * (LastPhraseInd - FirstPhraseInd + 1), 0) + FirstPhraseInd
                        
                        LastCategory = Category
                    Else
                        PhraseRow = PhraseRow + 1
                        If PhraseRow > LastPhraseInd Then
                            PhraseRow = FirstPhraseInd
                        End If
                    End If
                    
                    ' Get phrase from config and fill with content
                    Phrase = ActiveWorkbook.Worksheets("PhraseConfig").Range("Phrases").Cells(PhraseRow, 1).Value & " "
                    
                    Phrase = Replace(Phrase, Category, Content.Value)
                    
                    ' Fill in other config
                    While InStr(Phrase, "<") > 0
                        TagName = Mid(Phrase, InStr(Phrase, "<"), InStr(Phrase, ">") - InStr(Phrase, "<") + 1)
                        TagRow = Application.WorksheetFunction.Match(TagName, ActiveWorkbook.Worksheets("Variables").Range("Tags"), 0)
                        TagContent = ActiveWorkbook.Worksheets("Variables").Range("Tags").Cells(TagRow, 1).Offset(0, 1)
                        Phrase = Replace(Phrase, TagName, TagContent)
                    Wend
                    
                    ' Join into the paragraph
                    Paragraph = Paragraph & Phrase
                End If
            ' Otherwise, can directly load into template
            Else
                FindAndReplace WordApp.ActiveDocument, Tag.Value, Content.Value
            End If
        End If
        
    Next Tag
    
    If Not IsEmpty(Paragraph) Then
        FindAndReplace WordApp.ActiveDocument, Category, Left(Paragraph, Len(Paragraph) - 1)
    End If
    
    MsgBox Prompt:="Finish generating " & Direc & NewFile & ", please check.", Title:="Finish Generation"
End Sub

Private Sub FindAndReplace(Document, Find, Replace)
'
' FindAndReplace Sub
' Find characters in Word document and replace as specified string
' Params: Find = word to look for, Replace = word to replace with
'

'
    Dim ReplaceRemaining As String
    ' Execute replace only allows 250 char, so need to trim and use recursion to replace the rest
    If Len(Replace) > 240 Then
        ReplaceRemaining = Mid(Replace, 241)
        Replace = Left(Replace, 240) & Find
    End If
    
    Document.Content.Find.Execute _
            FindText:=Find, ReplaceWith:=Replace, Replace:=wdReplaceAll
    
    If Len(ReplaceRemaining) > 0 Then
        FindAndReplace Document, Find, ReplaceRemaining
    End If
End Sub

Private Function MatchLast(Lookupvalue As String, LookupRange As Range, ColumnNumber As Integer) As Long
    Dim i As Long
    For i = LookupRange.Columns(1).Cells.Count To 1 Step -1
        If Lookupvalue = LookupRange.Cells(i, 1) Then
            MatchLast = i
            Exit Function
        End If
    Next i
End Function

Private Function RangeExists(R As String) As Boolean
    Dim Test As Range
    On Error Resume Next
    Set Test = ActiveSheet.Range(R)
    RangeExists = (Err.Number = 0)
End Function

Private Function SheetExists(S As String, Optional Wb As Workbook) As Boolean
    Dim Test As Worksheet
    If Wb Is Nothing Then
        Wb = ActiveWorkbook
    End If
    On Error Resume Next
    Set Test = Wb.Sheets(S)
    SheetExists = (Not Test Is Nothing)
End Function
