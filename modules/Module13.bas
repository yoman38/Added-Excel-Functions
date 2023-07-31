Attribute VB_Name = "Module13"
' Function: lastCol
' Description: Returns the column number of the last value in a row, with the option to specify the sheet (optional).
' Parameters:
'   Optional sheetName: The name of the sheet to find the last column in. If not specified, the active sheet is used.
' Returns:
'   The column number of the last value in the row.
Function lastCol(Optional sheetName As String = "") As Long
    Dim ws As Worksheet
    Dim lastCell As Range
    
    If sheetName <> "" Then
        On Error Resume Next
        Set ws = Worksheets(sheetName)
        On Error GoTo 0
        
        If ws Is Nothing Then
            ' Sheet with the specified name not found.
            Exit Function
        End If
    Else
        Set ws = ActiveSheet
    End If
    
    On Error Resume Next
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False, _
        SearchFormat:=False)
    On Error GoTo 0
    
    If lastCell Is Nothing Then
        ' No value found in the sheet.
        lastCol = 0
    Else
        ' Return the column number of the last value.
        lastCol = lastCell.Column
    End If
End Function

' Function: lastRow
' Description: Returns the row number of the last value in a column, with the option to specify the sheet (optional).
' Parameters:
'   Optional sheetName: The name of the sheet to find the last row in. If not specified, the active sheet is used.
' Returns:
'   The row number of the last value in the column.
Function lastRow(Optional sheetName As String = "") As Long
    Dim ws As Worksheet
    Dim lastCell As Range
    
    If sheetName <> "" Then
        On Error Resume Next
        Set ws = Worksheets(sheetName)
        On Error GoTo 0
        
        If ws Is Nothing Then
            ' Sheet with the specified name not found.
            Exit Function
        End If
    Else
        Set ws = ActiveSheet
    End If
    
    On Error Resume Next
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, _
        SearchFormat:=False)
    On Error GoTo 0
    
    If lastCell Is Nothing Then
        ' No value found in the sheet.
        lastRow = 0
    Else
        ' Return the row number of the last value.
        lastRow = lastCell.row
    End If
End Function

' Function: lastUsedCol
' Description: Returns the column number of the last used column in the sheet.
' Parameters: None
' Returns:
'   The column number of the last used column in the sheet.
Function lastUsedCol() As Long
    On Error Resume Next
    lastUsedCol = ActiveSheet.Cells.Find(What:="*", After:=ActiveSheet.Cells(1, 1), LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False, _
        SearchFormat:=False).Column
    On Error GoTo 0
End Function

' Function: lastUsedRow
' Description: Returns the row number of the last used row in the sheet.
' Parameters: None
' Returns:
'   The row number of the last used row in the sheet.
Function lastUsedRow() As Long
    On Error Resume Next
    lastUsedRow = ActiveSheet.Cells.Find(What:="*", After:=ActiveSheet.Cells(1, 1), LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, _
        SearchFormat:=False).row
    On Error GoTo 0
End Function


