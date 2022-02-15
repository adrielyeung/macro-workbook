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

Sub Batch_ReplaceTagsWithContent()
'
' Batch_ReplaceTagsWithContent Sub
' Batch calling of Function ReplaceTagsWithContent(), finding for records in path FeedPath
' with PrepareStatus = "Prepare"
' And changing the basic config for each record
'

'
    Dim FeedPath As Range, PrepareStatus As Range, Status As Range, BatchConfig As Range, Config As Range
    Dim TemplateWb As Workbook, FeedWb As Workbook
    Dim i As Integer, SucCount As Integer, FailCount As Integer
    Dim OutMsg As String
    
    ' Set TemplateWb as the Excel holding the template config
    Set TemplateWb = ActiveWorkbook
    ' FeedPath is the path of the batch listings
    Set FeedPath = TemplateWb.Names("FeedPath").RefersToRange
    ' BatchConfig is the cells in the Template Excel to be updated for each batch run
    Set BatchConfig = TemplateWb.Names("BatchConfig").RefersToRange
    ' FeedWb is the batch listings Excel
    Set FeedWb = Workbooks.Open(FeedPath)
    ' PrepareStatus are the cells holding the processing status
    Set PrepareStatus = FeedWb.Names("PrepareStatus").RefersToRange
    ' Error msg from each individual run
    OutMsg = ""
    SucCount = 0
    FailCount = 0
    
    For Each Status In PrepareStatus
        ' Extract all "Prepare" statis
        If Status.Value = "Prepare" Then
            i = Status.Column - 1
            
            ' Copy each value in FeedWb to TemplateWb
            For Each Config In BatchConfig
                Config.Value = Status.Offset(0, -i).Value
                i = i - 1
                If i < 0 Then
                    MsgBox "There are more config cells than provided. Please check.", , "Batch failed"
                    Exit Sub
                End If
            Next Config
            TemplateWb.Activate
            ' Call the ReplaceTagsWithContent Function
            OutMsg = ReplaceTagsWithContent()
            If Left(OutMsg, 1) <> "!" Then
                Status.Value = "Check"
                SucCount = SucCount + 1
            Else
                FailCount = FailCount + 1
            End If
            Status.Offset(0, 1).Value = OutMsg
        ' Hit the end of file, finish
        ElseIf Status.Value = "End" Then
            Exit For
        End If
    Next Status
    
    If FailCount > 0 Then
        MsgBox Prompt:="Finish generating all pending records WITH ERROR, please check." & vbNewLine _
        & "Success count: " & SucCount & vbNewLine _
        & "Failed count: " & FailCount, Title:="ERROR in Generation"
    ElseIf SucCount = 0 Then
        MsgBox Prompt:="No pending records.", Title:="No Generation"
    Else
        MsgBox Prompt:="Finish generating all pending records SUCCESS." & vbNewLine _
        & "Count: " & SucCount, Title:="Finish Generation"
    End If
End Sub

Sub Single_ReplaceTagsWithContent()
'
' Single_ReplaceTagsWithContent Function
' Single calling of Function ReplaceTagsWithContent()
'

'
    Dim OutMsg As String
    
    OutMsg = ReplaceTagsWithContent()
    
    If Left(OutMsg, 1) = "!" Then
        MsgBox Prompt:="Finish generating WITH ERROR: " & Mid(OutMsg, 2), Title:="ERROR in Generation"
    Else
        MsgBox Prompt:="Finish generating SUCCESS, file location: " & OutMsg, Title:="Finish Generation"
    End If
End Sub

Private Function ReplaceTagsWithContent() As String
'
' ReplaceTagsWithContent Function
' Replace placeholder tags <xxx> with actual content according to Excel config
'

'
    Dim Tags As Range, Tag As Range, Content As Range, ListingPath As Range
    Dim Priority As Range, Order As Range, Field As Range, SearchField As Range, ListItem As Range
    Dim Missing As Boolean
    Dim RegEx As Object
    Dim Template As String, NewFile As String, Direc As String
    Dim CompName As String, ListItemString As String
    Dim Prefix As String, Suffix As String, Category As String, LastCategory As String, ListCategory As String
    Dim Phrase As String, Paragraph As String, TagName As String, TagContent As String, Listing As String
    Dim Random As Double, PhraseRow As Long, FirstInd As Long, LastInd As Long, TagRow As Integer, i As Long
    Dim ContentArr() As String
    
    Dim TemplateWb As Workbook, ListingWb As Workbook
    
    Dim WordApp As Word.Application
    Set WordApp = New Word.Application
    
    Set Tags = Range("Tags")
    
    ReplaceTagsWithContent = "!Error in program run, please check code"
    
    ' Open listing file if available
    Set TemplateWb = ActiveWorkbook
    Set ListingPath = TemplateWb.Names("ListingPath").RefersToRange
    If Not IsEmpty(ListingPath.Value) Then
        If Not Dir(ListingPath.Value) = "" Then
            Set ListingWb = Workbooks.Open(ListingPath)
        End If
    End If
    
    Missing = False
    
    TemplateWb.Activate
    
    ' Checking if any content field is missing
    For Each Tag In Tags
        
        If Not IsEmpty(Tag.Value) Then
            Set Content = Tag.Offset(0, 1)
            
            If Left(Tag.Value, 1) = "<" And IsEmpty(Content.Value) Then
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
        ReplaceTagsWithContent = "!Missing info in Generator config"
        Exit Function
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
        Prefix = "Document"
    End If
    
    If RangeExists("CompName") Then
        NewFile = Prefix & "_" & Range("CompName").Value & "_" & Format(Now(), "yyyymmdd_hhmmss") & ".docx"
    Else
        NewFile = Prefix & "_" & Format(Now(), "yyyymmdd_hhmmss") & ".docx"
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
    Listing = ""
    For Each Tag In Tags
        
        If Not IsEmpty(Tag.Value) Then
        
            Set Content = Tag.Offset(0, 1)
            Content.Value = Trim(Content.Value)
            ' If tag starts with "B", break into subitems and fill in corresponding "L" tags
            If Left(Tag.Value, 1) = "B" Then
                If Len(Content.Value) > 0 Then
                    Category = "L" & Mid(Tag.Value, 2)
                    
                    ContentArr = Split(Content.Value, ", ")
                    
                    FirstInd = Application.WorksheetFunction.Match(Category, Tags, 0)
                    LastInd = MatchLast(Category, Tags, 1)
                    
                    For i = FirstInd To LastInd
                        Range("Tags").Cells(i, 1).Offset(0, 1).Value = ""
                    Next i
                    
                    For i = FirstInd To LastInd
                        If i - FirstInd < UBound(ContentArr) - LBound(ContentArr) + 1 Then
                            Range("Tags").Cells(i, 1).Offset(0, 1).Value = Trim(ContentArr(i - FirstInd))
                        End If
                    Next i
                    
                    Category = ""
                    FirstInd = 0
                    LastInd = 0
                End If
            ' If tag starts with "P", loop phrases to load stock phrases
            ' If tag starts with "L", load from relevant records in listings file and insert
            ElseIf Left(Tag.Value, 1) = "P" Or Left(Tag.Value, 1) = "L" Then
                If Len(Content.Value) > 0 Then
                    ' Category = <xxx>
                    Category = Mid(Tag.Value, 2)
                    ' Changed category, add last category into template
                    If Category <> LastCategory And LastCategory <> "" Then
                        Randomize
                        Random = Rnd()
                        
                        WriteParagraph Paragraph, Listing, LastCategory, WordApp.ActiveDocument
                        Paragraph = ""
                        Phrase = ""
                        Listing = ""
                    End If
                    
                    ' Calculate new random index of stock phrases to start from
                    If Category <> LastCategory Then
                        FirstInd = Application.WorksheetFunction.Match(Category, ActiveWorkbook.Worksheets("PhraseConfig").Range("PhraseTags"), 0)
                        LastInd = MatchLast(Category, ActiveWorkbook.Worksheets("PhraseConfig").Range("PhraseTags"), 1)
                        
                        PhraseRow = Application.WorksheetFunction.RoundDown(Random * (LastInd - FirstInd + 1), 0) + FirstInd
                        
                        LastCategory = Category
                    Else
                        PhraseRow = PhraseRow + 1
                        If PhraseRow > LastInd Then
                            PhraseRow = FirstInd
                        End If
                    End If
                    
                    ' Get phrase from config and fill with content
                    Phrase = ActiveWorkbook.Worksheets("PhraseConfig").Range("Phrases").Cells(PhraseRow, 1).Value
                    
                    If Left(Tag.Value, 1) = "P" Then
                        Phrase = Replace(Phrase, Category, Content.Value)
                    ' Fill in list item from listing file
                    ElseIf Left(Tag.Value, 1) = "L" Then
                        Phrase = Replace(Phrase, Right(Tag.Value, Len(Tag.Value) - 1), Content.Value)
                        Listing = Listing & Content.Value & ", "
                        If ListingWb Is Nothing Then
                            Paragraph = Range("ListItem").Cells(FirstInd, 1).Value
                        Else
                            ListingWb.Activate
                            Set Priority = ListingWb.Names("Priority").RefersToRange
                            Set Order = ListingWb.Names("Order").RefersToRange
                            
                            Priority.Worksheet.AutoFilter.Sort.SortFields.Clear
                            Priority.Worksheet.AutoFilter.Sort.SortFields.Add2 Key:=Priority, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
                                xlSortNormal
                            With Priority.Worksheet.AutoFilter.Sort
                                .Header = xlYes
                                .MatchCase = False
                                .Orientation = xlTopToBottom
                                .SortMethod = xlPinYin
                                .Apply
                            End With
                            
                            Set SearchField = ListingWb.Names("SearchField").RefersToRange
                            Set ListItem = ListingWb.Names("ListItem").RefersToRange
                            
                            For Each Field In SearchField.Cells
                                If InStr(1, Field.Value, Content.Value) > 0 Then
                                    ListItemString = Trim(ListItem.Cells(Field.Row, 1).Value)
                                    
                                    Set RegEx = New RegExp
                                    
                                    RegEx.Pattern = "-+\s"
                                    RegEx.Global = True
                                    
                                    ListItemString = RegEx.Replace(ListItemString, "")
                                    
                                    ListItemString = StrConv(Left(ListItemString, 1), vbLowerCase) & Right(ListItemString, Len(ListItemString) - 1)
                                    
                                    While InStr(1, ListItemString, vbLf) > 0
                                        Phrase = Replace(Phrase, "<ListItem>", Left(ListItemString, InStr(1, ListItemString, vbLf) - 1) & ", <ListItem>")
                                        ListItemString = Mid(ListItemString, InStr(1, ListItemString, vbLf) + 1)
                                        ListItemString = StrConv(Left(ListItemString, 1), vbLowerCase) & Right(ListItemString, Len(ListItemString) - 1)
                                    Wend
                                    
                                    Phrase = Replace(Phrase, "<ListItem>", ListItemString)
                                    Exit For
                                End If
                            Next Field
                            
                            Order.Worksheet.AutoFilter.Sort.SortFields.Clear
                            Order.Worksheet.AutoFilter.Sort.SortFields.Add2 Key:=Order, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
                                xlSortNormal
                            With Order.Worksheet.AutoFilter.Sort
                                .Header = xlYes
                                .MatchCase = False
                                .Orientation = xlTopToBottom
                                .SortMethod = xlPinYin
                                .Apply
                            End With
                            
                            TemplateWb.Activate
                        End If
                    End If
                    
                    ' Fill in other config
                    While InStr(Phrase, "<") > 0
                        TagName = Mid(Phrase, InStr(Phrase, "<"), InStr(Phrase, ">") - InStr(Phrase, "<") + 1)
                        TagRow = Application.WorksheetFunction.Match(TagName, ActiveWorkbook.Worksheets("Variables").Range("Tags"), 0)
                        TagContent = ActiveWorkbook.Worksheets("Variables").Range("Tags").Cells(TagRow, 1).Offset(0, 1)
                        Phrase = Replace(Phrase, TagName, TagContent)
                    Wend
                    
                    ' Join into the paragraph
                    Paragraph = Paragraph & Phrase & " "
                End If
            ' Otherwise, can directly load into template
            Else
                If Len(Paragraph) > 0 Then
                    WriteParagraph Paragraph, Listing, Category, WordApp.ActiveDocument
                    
                    Paragraph = ""
                    Phrase = ""
                    Listing = ""
                End If
                LastCategory = ""
                FindAndReplace WordApp.ActiveDocument, Tag.Value, Content.Value
            End If
        End If
    Next Tag
    
    If Len(Paragraph) > 0 Then
        WriteParagraph Paragraph, Listing, Category, WordApp.ActiveDocument
    End If
    
    ReplaceTagsWithContent = Direc & NewFile
End Function

Private Function AddAndToList(Str As String) As String
' AddAndToList Function
' For a comma-separated String, find last comma and add "and" after it
' e.g. "a, b, c" -> "a, b, and c"
    
    Dim LastPos As Long
    LastPos = LastPosInString(Str, ",")
    
    If LastPos = 0 Then
        AddAndToList = Str
        Exit Function
    End If
    
    AddAndToList = Left(Str, LastPos + 1) & "and " & Mid(Str, LastPos + 2)
End Function

Private Function LastPosInString(Str As String, Ch As String) As Long
' LastPosInString Function
' Find the index of last occurence of Ch in Str, return 0 if none found
'
    For LastPosInString = Len(Str) To 1 Step -1
        If Mid(Str, LastPosInString, 1) = Ch Then
            Exit Function
        End If
    Next LastPosInString
    
    LastPosInString = 0
End Function

Private Sub WriteParagraph(Paragraph As String, Listing As String, Category As String, WordDocument As Document)
' WriteParagraph Sub
' Join Listing (main points) and Paragraph and fill in Word document
'

    ' Remove last comma and space of Listing then join into Paragraph
    If Len(Listing) > 0 Then
        Listing = AddAndToList(Left(Listing, Len(Listing) - 2)) & ". "
        Paragraph = Listing & Paragraph
    End If
    
    ' Remove the last space of Paragraph before inserting into template
    FindAndReplace WordDocument, Category, Left(Paragraph, Len(Paragraph) - 1)
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

Private Function RangeExists(R As String) As Boolean
' RangeExists Function
' Test if Range R exists
'
    Dim Test As Range
    On Error Resume Next
    Set Test = ActiveSheet.Range(R)
    RangeExists = (Err.Number = 0)
End Function

Private Function SheetExists(S As String, Optional Wb As Workbook) As Boolean
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
