Attribute VB_Name = "Module2"
' Function: EQUIV_X
' Description: Searches for a value in a range of cells and returns its position.
' Parameters:
'   searchValue: The value to be searched in the range.
'   searchRange: The range of cells where the search will be performed.
' Returns:
'   If the value is found, it returns the relative position of the value in the range (1-based index).
'   If the value is not found, it returns an error value (e.g., #N/A).
Function EQUIV_X(searchValue As Variant, searchRange As Range) As Variant
    Dim result As Variant
    On Error Resume Next
    ' Using the Match function to find the position of the value in the range.
    result = Application.WorksheetFunction.match(searchValue, searchRange, 0)
    On Error GoTo 0
    
    ' If the value is found, return the position, else return an error value (#N/A).
    If Not IsError(result) Then
        EQUIV_X = result
    Else
        EQUIV_X = CVErr(xlErrNA)
    End If
End Function

' Function: SEARCH_X
' Description: Searches for a value in a range of cells and returns a value at the same position in another range of cells.
' Parameters:
'   searchValue: The value to be searched in the searchRange.
'   searchRange: The range of cells where the search will be performed.
'   returnRange: The range of cells from which the corresponding value will be returned.
' Returns:
'   If the value is found, it returns the value at the same position in the returnRange.
'   If the value is not found, it returns an error value (e.g., #N/A).
Function SEARCH_X(searchValue As Variant, searchRange As Range, returnRange As Range) As Variant
    Dim position As Variant
    On Error Resume Next
    ' Using the Match function to find the position of the value in the searchRange.
    position = Application.WorksheetFunction.match(searchValue, searchRange, 0)
    On Error GoTo 0
    
    ' If the value is found, return the corresponding value from the returnRange.
    ' Otherwise, return an error value (#N/A).
    If Not IsError(position) Then
        SEARCH_X = returnRange.Cells(position).value
    Else
        SEARCH_X = CVErr(xlErrNA)
    End If
End Function

' Function: IF_NON_EMPTY
' Description: Checks if a cell is not empty and returns a result based on the test.
' Parameters:
'   CheckCell: The cell to check if it's not empty.
'   IfEmpty: The value to return if the cell is empty.
'   IfNotEmpty: The value to return if the cell is not empty.
' Returns:
'   IfEmpty if the cell is empty, or IfNotEmpty if the cell is not empty.
Function IF_NON_EMPTY(CheckCell As Range, IfEmpty As Variant, IfNotEmpty As Variant) As Variant
    ' Check if the cell is empty.
    If CheckCell.value = "" Then
        ' Return IfEmpty if the cell is empty.
        IF_NON_EMPTY = IfEmpty
    Else
        ' Return IfNotEmpty if the cell is not empty.
        IF_NON_EMPTY = IfNotEmpty
    End If
End Function

' Function: SHEET_NAME
' Description: Returns the name of the worksheet for a given cell.
' Parameters:
'   Cell: The cell from which to get the worksheet name.
' Returns:
'   The name of the worksheet where the cell is located.
Function SHEET_NAME(cell As Range) As String
    ' Use the Parent property to get the worksheet containing the cell.
    ' Then, use the Name property to get the name of the worksheet.
    SHEET_NAME = cell.Parent.Name
End Function

