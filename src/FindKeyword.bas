Attribute VB_Name = "Module1"
Sub FindKeyword()
Attribute FindKeyword.VB_Description = "Find a number of keywords within range and display matched keywords in a new column"
Attribute FindKeyword.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Find_keyword Macro
' Find a number of keywords within range and display matched keywords in a new column
'

'
    Dim InputFileName As String, SearchWordArray() As String, SearchRangeString As String, SearchRange As Range, ResultColumnString As String
    InputFileName = InputBox("Please enter full file path with list of search words", "Input File", "C:\Users\Adriel\OneDrive\¤å¥ó\PythonScripts\Scraping\Report\search_word.txt")
    If StrPtr(InputFileName) = 0 Or StrComp(InputFileName, "") = 0 Then
        MsgBox Prompt:="Process cancelled.", Title:="Abort"
        Exit Sub
    End If
    
    SearchRangeString = InputBox("Please enter search range, separated by colon ':' (e.g. 'A1:B3', 'C:C')", "Search Range", "E:E")
    If StrPtr(SearchRangeString) = 0 Or StrComp(SearchRangeString, "") = 0 Then
        MsgBox Prompt:="Process cancelled.", Title:="Abort"
        Exit Sub
    End If
    
    ResultColumnString = InputBox("Please enter the column you would like search results to be in (e.g. 'K')", "Result Column", "I")
    If StrPtr(ResultColumnString) = 0 Or StrComp(ResultColumnString, "") = 0 Then
        MsgBox Prompt:="Process cancelled.", Title:="Abort"
        Exit Sub
    End If
    
    Set SearchRange = ActiveSheet.Range(SearchRangeString)
    
    ' Clear result column
    Range(ResultColumnString & ":" & ResultColumnString).Clear
    
    Open InputFileName For Input As #1
        SearchWordArray = Split(Input(LOF(1), #1), ",")
    Close #1
    
    Dim RowResultList() As Integer, RowResultListUnique(), UniqueResultDict As Object, SearchWord, SearchWordTrimmed As String, Row, r As Integer
    Set UniqueResultDict = CreateObject("Scripting.Dictionary")
    
    For Each SearchWord In SearchWordArray
    
        SearchWordTrimmed = Replace(Replace(WorksheetFunction.Trim((SearchWord)), vbLf, ""), vbCr, "")
    
        RowResultList = FindAll(SearchRange, SearchWordTrimmed)
        
        For r = LBound(RowResultList) To UBound(RowResultList)
            UniqueResultDict(RowResultList(r)) = Empty
        Next
        
        RowResultListUnique = UniqueResultDict.keys()
        
        For Each Row In RowResultListUnique
            Range(ResultColumnString & (Row)).Value = Range(ResultColumnString & (Row)).Value & SearchWordTrimmed & ", "
        Next
    Next
    
    With ActiveSheet.Range(ResultColumnString & "1").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With ActiveSheet.Range(ResultColumnString & "1").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With ActiveSheet.Range(ResultColumnString & "1").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With ActiveSheet.Range(ResultColumnString & "1").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With ActiveSheet.Range(ResultColumnString & "1")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Value = "Keywords"
    End With
    
    ActiveSheet.Columns(ResultColumnString & ":" & ResultColumnString).EntireColumn.AutoFit
    If Not ActiveSheet.AutoFilterMode Then
        ActiveSheet.Range("A1").AutoFilter
    End If
    
End Sub

Function FindAll(SearchRange As Range, SearchString As String) As Variant
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

