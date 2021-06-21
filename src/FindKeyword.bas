Attribute VB_Name = "Module1"
Sub FindKeyword()
Attribute FindKeyword.VB_Description = "Find a number of keywords within range and display matched keywords in a new column"
Attribute FindKeyword.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Find_keyword Macro
' Find a number of keywords within range and display matched keywords in a new column
'

'
    Dim InputFileName As String, SearchWordArray() As String, SearchRangeString As String, SearchRange As Range, ResultColumnString As String, ScoreColumnString As String, ScoreMode As VbMsgBoxResult
    Dim Score As Integer
    InputFileName = InputBox("Please enter full file path with list of search words", "Input File", "C:\Users\Adriel\OneDrive\¤å¥ó\PythonScripts\Scraping\Report\search_word.txt")
    If StrPtr(InputFileName) = 0 Or StrComp(InputFileName, "") = 0 Then
        MsgBox Prompt:="Process cancelled.", Title:="Abort"
        Exit Sub
    End If
    
    SearchRangeString = InputBox("Please enter search range, separated by colon ':' (e.g. 'A1:B3', 'C:D')", "Search Range", "E:F")
    If StrPtr(SearchRangeString) = 0 Or StrComp(SearchRangeString, "") = 0 Then
        MsgBox Prompt:="Process cancelled.", Title:="Abort"
        Exit Sub
    End If
    
    ResultColumnString = InputBox("Please enter the column you would like search results to be in (e.g. 'K')", "Result Column", "I")
    If StrPtr(ResultColumnString) = 0 Or StrComp(ResultColumnString, "") = 0 Then
        MsgBox Prompt:="Process cancelled.", Title:="Abort"
        Exit Sub
    End If
    ScoreColumnString = ShiftToNextColumn(ResultColumnString)

    Set SearchRange = ActiveSheet.Range(SearchRangeString)
    
    ' Clear result column
    Range(ResultColumnString & ":" & ResultColumnString).Clear
    
    Open InputFileName For Input As #1
        SearchWordArray = Split(Input(LOF(1), #1), ",")
    Close #1
    
    ScoreMode = MsgBox("Do you want to add weights to your search word list?" & vbNewLine & _
        "Select Yes (Y) - Highest score (" & UBound(SearchWordArray) - LBound(SearchWordArray) + 1 & " points) on the first keyword '" & SearchWordArray(0) & "', then 1 less point on each successive word." & vbNewLine & _
        "Select No (N) - All search keywords will have 1 point.", vbYesNo, "Add weights to search words")
        
    ' Find score of each word
    If ScoreMode = vbYes Then
        Score = UBound(SearchWordArray) - LBound(SearchWordArray) + 1
    ElseIf StrComp(ScoreMode, "Equal - all words in search word list have equal scores") Then
        Score = 1
    End If
    
    Dim RowResultList() As Integer, RowResultListUnique(), UniqueResultDict As Object, SearchWord, SearchWordTrimmed As String, Row, r As Integer
    
    Range(ScoreColumnString & "2:" & ScoreColumnString & ActiveSheet.UsedRange.Rows.Count).Value = 0
    
    ' Add search keywords into keywords field
    For Each SearchWord In SearchWordArray
    
        Set UniqueResultDict = CreateObject("Scripting.Dictionary")
    
        SearchWordTrimmed = Replace(Replace(WorksheetFunction.Trim((SearchWord)), vbLf, ""), vbCr, "")
    
        RowResultList = FindAll(SearchRange, SearchWordTrimmed)
        
        For r = LBound(RowResultList) To UBound(RowResultList)
            UniqueResultDict(RowResultList(r)) = Empty
        Next
        
        RowResultListUnique = UniqueResultDict.keys()
        
        For Each Row In RowResultListUnique
            Range(ResultColumnString & (Row)).Value = Range(ResultColumnString & (Row)).Value & SearchWordTrimmed & ", "
            
            Range(ScoreColumnString & (Row)).Value = Range(ScoreColumnString & (Row)).Value + Score
        Next
        
        If ScoreMode = vbYes Then
            Score = Score - 1
        End If
        
    Next
    
    SetHeaderFormat ResultColumnString, "Keywords"
    SetHeaderFormat ScoreColumnString, "Score"
    
    ActiveSheet.Columns(ResultColumnString & ":" & ResultColumnString).EntireColumn.AutoFit
    If Not ActiveSheet.AutoFilterMode Then
        ActiveSheet.Range("A1").AutoFilter
    Else
        ActiveSheet.Range("A1").AutoFilter
        ActiveSheet.Range("A1").AutoFilter
    End If
    
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With
    
    ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveSheet.AutoFilter.Sort.SortFields.Add2 Key:= _
        ActiveSheet.UsedRange.Columns(ScoreColumnString), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
End Sub

Private Sub SetHeaderFormat(Column As String, Header As String)
'
' SetHeaderFormat sub
' Initialises a new column as in 'Column' with its heading on the 1st row
'

    With ActiveSheet.Range(Column & "1").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With ActiveSheet.Range(Column & "1").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With ActiveSheet.Range(Column & "1").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With ActiveSheet.Range(Column & "1").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With ActiveSheet.Range(Column & "1")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Value = Header
    End With

End Sub

Private Function ShiftToNextColumn(ColumnString As String)
'
' ShiftToNextColumn function
' Find the next column address from previous
'

    ShiftToNextColumn = Replace(Replace(Range(ColumnString & "1").Offset(0, 1).Address, "1", ""), "$", "")

End Function

Private Function FindAll(SearchRange As Range, SearchString As String) As Variant
'
' FindAll function
' Find all cells matching SearchString within SearchRange, returning the row number of matches.
'

'

    Dim FoundCell As Range, FirstFound As String, Output() As Integer, i As Integer
    
    Set FoundCell = SearchRange.Find(What:=SearchString, LookIn:=xlValues, LookAt:=xlPart, _
                SearchOrder:=xlRows, SearchDirection:=xlNext, MatchCase:=False, MatchByte:=True, SearchFormat:=False)
                
    If FoundCell Is Nothing Then
        GoTo NothingFound
    End If
    
    FirstFound = FoundCell.Address
    ReDim Preserve Output(0)
    Output(0) = FoundCell.Row
    i = 1
    
    Do Until FoundCell Is Nothing
        Set FoundCell = SearchRange.FindNext(After:=FoundCell)
        
        If StrComp(FoundCell.Address, FirstFound) = 0 Then Exit Do
        
        ReDim Preserve Output(i)
        Output(i) = FoundCell.Row

        i = i + 1
    Loop
    
    FindAll = Output
    
Exit Function
    
NothingFound:
    Set FindAll = Nothing

End Function

