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

Sub Batch_ReplaceTagsWithContent_JobList()
'
' Batch_ReplaceTagsWithContent_JobList Sub
' Batch calling of Function ReplaceTagsWithContent(), finding for records in FeedWb workbook
' Called from JobListing Excel (FeedWb)
'

'
    Dim GeneratorPath As Range
    Dim FeedWb As Workbook, GeneratorWb As Workbook
    
    ' FeedWb is the batch listings Excel
    Set FeedWb = ActiveWorkbook
    ' GeneratorPath is the path of the Generator Excel
    Set GeneratorPath = FeedWb.Names("GeneratorPath").RefersToRange
    
    ' Set GeneratorWb as the Generator Excel
    Set GeneratorWb = Workbooks.Open(GeneratorPath)
    
    Batch_ReplaceTagsWithContent FeedWb, GeneratorWb

End Sub

Sub Batch_ReplaceTagsWithContent_Generator()
'
' Batch_ReplaceTagsWithContent Sub
' Batch calling of Function ReplaceTagsWithContent(), finding for records in FeedWb workbook
' Called from Generator Excel (GeneratorWb)
'

'
    Dim FeedPath As Range
    Dim FeedWb As Workbook, GeneratorWb As Workbook
    
    ' Set GeneratorWb as the Excel holding the template config
    Set GeneratorWb = ActiveWorkbook
    ' FeedPath is the path of the batch listings
    Set FeedPath = GeneratorWb.Names("FeedPath").RefersToRange
    ' FeedWb is the batch listings Excel
    Set FeedWb = Workbooks.Open(FeedPath)
    
    Batch_ReplaceTagsWithContent FeedWb, GeneratorWb
    
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

Private Sub Batch_ReplaceTagsWithContent(FeedWb As Workbook, GeneratorWb As Workbook)
'
' Batch_ReplaceTagsWithContent Sub
' Batch calling of Function ReplaceTagsWithContent()
' with PrepareStatus = "Prepare"
' And changing the basic config for each record
'

'
    Dim PrepareStatus As Range, Status As Range, BatchConfig As Range, Config As Range
    Dim i As Integer, SucCount As Integer, FailCount As Integer
    Dim OutMsg As String
    
    ' BatchConfig is the cells in the Template Excel to be updated for each batch run
    Set BatchConfig = GeneratorWb.Names("BatchConfig").RefersToRange
    
    ' PrepareStatus are the cells holding the processing status
    Set PrepareStatus = FeedWb.Names("PrepareStatus").RefersToRange
    ' Error msg from each individual run
    OutMsg = ""
    SucCount = 0
    FailCount = 0
    
    For Each Status In PrepareStatus
        ' Extract all "Prepare" status
        If Status.Value = "Prepare" Then
            i = Status.Column - 1
            
            ' Copy each value in FeedWb to GeneratorWb
            For Each Config In BatchConfig
                Config.Value = Status.Offset(0, -i).Value
                i = i - 1
                If i < 0 Then
                    MsgBox "There are more config cells than provided. Please check.", , "Batch failed"
                    GoTo Output
                End If
            Next Config
            GeneratorWb.Activate
            ' Call the ReplaceTagsWithContent Function
            OutMsg = ReplaceTagsWithContent()
            If Left(OutMsg, 1) <> "!" Then
                Status.Value = "GenPDF"
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
    
Output:
    If FailCount > 0 Then
        MsgBox Prompt:="Finish generating all pending records WITH ERROR, please check." & vbNewLine _
        & "Success count: " & SucCount & vbNewLine _
        & "Failed count: " & FailCount, Title:="ERROR in Generation"
    ElseIf SucCount = 0 Then
        MsgBox Prompt:="No records generated.", Title:="No Generation"
    Else
        MsgBox Prompt:="Finish generating all pending records SUCCESS." & vbNewLine _
        & "Count: " & SucCount, Title:="Finish Generation"
    End If
End Sub

Private Function ReplaceTagsWithContent() As String
'
' ReplaceTagsWithContent Function
' Replace placeholder tags <xxx> with actual content according to Excel config
'

'
    Dim Tags As Range, Tag As Range, Content As Range, ListingPath As Range
    Dim Priority As Range, Order As Range, Field As Range, SearchField As Range
    Dim ListDesc As Range, ListJobTitle As Range, ListComp As Range
    Dim Missing As Boolean
    Dim RegEx As Object
    Dim Template As String, NewFile As String, Direc As String, CompName As String
    Dim ListDescString As String, ListJobTitleString As String, ListCompString As String, PrevJobTitleString As String, PrevCompString As String
    Dim Prefix As String, Suffix As String, Category As String, LastCategory As String, ListCategory As String
    Dim Phrase As String, Paragraph As String, TagName As String, TagContent As String, Listing As String
    Dim ListDescTag As String, ListJobTitleTag As String, ListCompTag As String
    Dim Random As Double, PhraseRow As Long, FirstInd As Long, LastInd As Long, TagRow As Integer, i As Long, LPDiff As Integer
    Dim ContentArr() As String
    
    Dim GeneratorWb As Workbook, ListingWb As Workbook
    
    Dim WordApp As Word.Application
    Set WordApp = New Word.Application
    
    Set Tags = Range("Tags")
    Set RegEx = New RegExp
    
    ListDescTag = "<ListDesc>"
    ListCompTag = "<ListComp>"
    ListJobTitleTag = "<ListJobTitle>"
    
    PrevCompTag = "<PrevComp>"
    PrevJobTitleTag = "<PrevJobTitle>"
    
    ' "!" in first position for an error
    ReplaceTagsWithContent = "!Error in program run, please check code"
    
    ' Open listing file if available
    Set GeneratorWb = ActiveWorkbook
    Set ListingPath = GeneratorWb.Names("ListingPath").RefersToRange
    If Not IsEmpty(ListingPath.Value) Then
        If Not Dir(ListingPath.Value) = "" Then
            Set ListingWb = Workbooks.Open(ListingPath)
        End If
    End If
    
    Missing = False
    
    GeneratorWb.Activate
    
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
        ReplaceTagsWithContent = "!Missing highlighted fields in Generator config, please check."
        Exit Function
    End If
    
    ' Set up template file
    If ExcelUtil.RangeExists("Template") Then
        Template = Range("Template").Value
    Else
        Template = InputBox("Please enter template file path", "Template")
    End If
    Direc = Left(Template, InStrRev(Template, Application.PathSeparator))
    
    ' Duplicate template file for filling in
    If ExcelUtil.RangeExists("FileNamePrefix") Then
        Prefix = Range("FileNamePrefix")
    Else
        Prefix = "Document"
    End If
    
    If ExcelUtil.RangeExists("CompName") Then
        NewFile = Prefix & "_" & Range("CompName").Value & "_" & Format(Now(), "yyyymmdd_hhmmss") & ".docx"
    Else
        NewFile = Prefix & "_" & Format(Now(), "yyyymmdd_hhmmss") & ".docx"
    End If

OpenFile:
    On Error GoTo FileError:
    FileCopy Template, Direc & NewFile
    
    ChDir Direc

    WordApp.Documents.Open Direc & NewFile
    On Error GoTo 0
    WordApp.Visible = True
    
    ' Find and replace tags in Word template
    If ExcelUtil.SheetExists("DateConfig", ActiveWorkbook) Then
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
                    
                    ' Convert all newlines into commas, then separate into list
                    ContentArr = Split(Replace(Content.Value, vbLf, ", "), ", ")
                    
                    FirstInd = Application.WorksheetFunction.Match(Category, Tags, 0)
                    LastInd = ExcelUtil.MatchLast(Category, Tags, 1)
                    
                    For i = FirstInd To LastInd
                        Tags.Cells(i, 1).Offset(0, 1).Value = ""
                    Next i
                    
                    For i = FirstInd To LastInd
                        If i - FirstInd < UBound(ContentArr) - LBound(ContentArr) + 1 Then
                            Tags.Cells(i, 1).Offset(0, 1).Value = Trim(ContentArr(i - FirstInd))
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
                        LastInd = ExcelUtil.MatchLast(Category, ActiveWorkbook.Worksheets("PhraseConfig").Range("PhraseTags"), 1)
                        
                        PhraseRow = Application.WorksheetFunction.RoundDown(Random * (LastInd - FirstInd + 1), 0) + FirstInd
                        
                        ' Find the difference between the replacement "P" item from the "L" item in Generator config,
                        ' in case search in List config failed
                        If Left(Tag.Value, 1) = "L" Then
                            LPDiff = Application.WorksheetFunction.Match("P" & Range("ListItem").Cells(FirstInd, 1).Value, Tags, 0) + Tags.Row - 1 - Tag.Row
                        End If
                        
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
                            ' Search in ListingWb for relevant items
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
                            
                            ' Search in ListingWb, SearchField
                            Set SearchField = ListingWb.Names("SearchField").RefersToRange
                            Set ListDesc = ListingWb.Names("ListDesc").RefersToRange
                            Set ListComp = ListingWb.Names("ListComp").RefersToRange
                            Set ListJobTitle = ListingWb.Names("ListJobTitle").RefersToRange
                            
                            For Each Field In SearchField.Cells
                                If InStr(1, Field.Value, Content.Value) > 0 Then
                                    ListDescString = Trim(ListDesc.Cells(Field.Row, 1).Value)
                                    ListCompString = Trim(ListComp.Cells(Field.Row, 1).Value)
                                    ListJobTitleString = Trim(ListJobTitle.Cells(Field.Row, 1).Value)
                                                                        
                                    RegEx.Pattern = "-+\s"
                                    RegEx.Global = True
                                    
                                    ' Remove all bullet points
                                    ListDescString = RegEx.Replace(ListDescString, "")
                                    
                                    ListDescString = StrConv(Left(ListDescString, 1), vbLowerCase) & Right(ListDescString, Len(ListDescString) - 1)
                                    
                                    ' Join up multiple items
                                    While InStr(1, ListDescString, vbLf) > 0
                                        Phrase = Replace(Phrase, ListDescTag, Left(ListDescString, InStr(1, ListDescString, vbLf) - 1) & ", <ListDesc>")
                                        ListDescString = Mid(ListDescString, InStr(1, ListDescString, vbLf) + 1)
                                        ListDescString = StrConv(Left(ListDescString, 1), vbLowerCase) & Right(ListDescString, Len(ListDescString) - 1)
                                    Wend
                                    
                                    Phrase = Replace(Phrase, ListDescTag, ListDescString)
                                    Phrase = Replace(Phrase, ListCompTag, ListCompString)
                                    Phrase = Replace(Phrase, ListJobTitleTag, ListJobTitleString)
                                    Exit For
                                End If
                            Next Field
                            
                            ' If above replace failed, find relevant item in Generator config to replace
                            If InStr(1, Phrase, ListDescTag) > 0 Then
                                Phrase = Replace(Phrase, ListDescTag, Tag.Offset(LPDiff, 1).Value)
                                ' Replace ListCompTag and ListJobTitleTag with previous company and previous job title tags
                                Phrase = Replace(Phrase, ListCompTag, PrevCompTag)
                                Phrase = Replace(Phrase, ListJobTitleTag, PrevJobTitleTag)
                            End If
                            
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
                            
                            GeneratorWb.Activate
                        End If
                    End If
                    
                    ' Fill in other config
                    While InStr(Phrase, "<") > 0
                        TagName = Mid(Phrase, InStr(Phrase, "<"), InStr(Phrase, ">") - InStr(Phrase, "<") + 1)
                        TagRow = Application.WorksheetFunction.Match(TagName, Tags, 0)
                        TagContent = Tags.Cells(TagRow, 1).Offset(0, 1)
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
    Exit Function
    
FileError:
    Do While True
        If MsgBox("Please close all Word Documents for the program to continue." & vbNewLine & _
            "Click [OK] to continue. Click [Cancel] to skip generation for this file." & vbNewLine & _
            Path, vbOKCancel, "Word Documents Open") = vbCancel Then
            Exit Function
        Else
            GoTo OpenFile
        End If
    Loop
End Function

Private Function AddAndToList(Str As String) As String
' AddAndToList Function
' For a comma-separated String, find last comma and add "and" after it
' e.g. "a, b, c" -> "a, b, and c"
    
    Dim LastPos As Long
    LastPos = InStrRev(Str, ",")
    
    If LastPos = 0 Then
        AddAndToList = Str
        Exit Function
    End If
    
    AddAndToList = Left(Str, LastPos + 1) & "and " & Mid(Str, LastPos + 2)
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
