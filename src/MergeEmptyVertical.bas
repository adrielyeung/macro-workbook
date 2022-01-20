Attribute VB_Name = "MergeEmptyVertical"
Sub MergeEmptyVertical()
Attribute MergeEmptyVertical.VB_Description = "Merge selected rows for each column (but do not center)."
Attribute MergeEmptyVertical.VB_ProcData.VB_Invoke_Func = " \n14"
'
' MergeEmptyVertical Macro
' Merge empty cells vertically for each column (but do not center).
' For each merge, count number of cells merged and output to designated column.
'

'
    Dim ColInd As Long, RowInd As Long, ColStart As Long, ColEnd As Long, RowStart As Long, RowEnd As Long, MergeCount As Long
    Dim StartCell As String, EndCell As String, OutputColumn As String
    Dim FirstMerge As Boolean
    
    ' Output column for number of rows merged
    OutputColumn = Application.InputBox(Title:="Output column", Prompt:="Please provide merged column count output column (e.g. A)?", Type:=2, Default:="H")
    
    ColStart = Selection.Columns(1).Column
    ColEnd = Selection.Columns.Count + ColStart - 1
    RowStart = Selection.Rows(1).Row
    RowEnd = Selection.Rows.Count + RowStart - 1
    
    ' For each column in used range
    For ColInd = ColStart To ColEnd
        ' First time merge, no previous merge
        FirstMerge = True
        MergeCount = 0
        StartCell = Cells(RowStart, ColInd).Address()
        For RowInd = RowStart To RowEnd
            ' Find next non-empty cell
            If Not IsEmpty(Cells(RowInd, ColInd)) Then
                ' Finish previous merge
                If Not FirstMerge Then
                    EndCell = Cells(RowInd - 1, ColInd).Address()
                    ' Print merged column from start to end
                    If Len(OutputColumn) = 1 Then
                        Cells(Range(StartCell).Row, Range(OutputColumn & 1).Column) = MergeCount
                    End If
                    Range(StartCell, EndCell).Merge
                    MergeCount = 0
                End If
                FirstMerge = False
                StartCell = Cells(RowInd, ColInd).Address()
            End If
            MergeCount = MergeCount + 1
        Next
        ' Return to last row index
        RowInd = RowInd - 1
        If Len(OutputColumn) = 1 Then
            Cells(Range(StartCell).Row, Range(OutputColumn & 1).Column) = MergeCount
        End If
        EndCell = Cells(RowEnd, ColInd).Address()
        Range(StartCell, EndCell).Merge
        MergeCount = 0
    Next
End Sub
