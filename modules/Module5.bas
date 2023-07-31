Attribute VB_Name = "Module5"
' Function: JOIN
' Description: Joins the values of a cell range using a specified separator (optional).
' Parameters:
'   SourceRange: The range of cells to join values from.
'   Optional Separator: The separator character used to join the values (default is an empty string).
' Returns:
'   The joined string of cell values.
Function JOIN(SourceRange As Range, Optional separator As String = "") As String
    Dim cell As Range
    Dim result As String
    
    result = ""
    
    ' Loop through each cell in the SourceRange.
    For Each cell In SourceRange
        ' Append the cell value to the result string, separated by the separator.
        result = result & cell.value & separator
    Next cell
    
    ' Remove the trailing separator if it exists.
    If Len(separator) > 0 Then
        result = Left(result, Len(result) - Len(separator))
    End If
    
    ' Return the joined string of cell values.
    JOIN = result
End Function

' Function: JOIN_SORT
' Description: Joins the values of a cell range using a specified separator (optional) and in ascending (default) or descending order.
' Parameters:
'   SourceRange: The range of cells to join values from.
'   Optional Separator: The separator character used to join the values (default is an empty string).
'   Optional SortOrder: The sorting order of the joined values ("ASC" for ascending, "DESC" for descending).
' Returns:
'   The joined string of cell values in the specified order.
Function JOIN_SORT(SourceRange As Range, Optional separator As String = "", Optional SortOrder As String = "ASC") As String
    Dim cell As Range
    Dim result As String
    Dim arr() As String
    
    result = ""
    
    ' Loop through each cell in the SourceRange and store the values in an array.
    For Each cell In SourceRange
        ReDim Preserve arr(0 To UBound(arr) + 1)
        arr(UBound(arr)) = cell.value
    Next cell
    
    ' Sort the array based on the specified SortOrder.
    If UCase(SortOrder) = "DESC" Then
        SortArrayDescending arr
    Else
        SortArrayAscending arr
    End If
    
    ' Join the sorted array with the separator.
    result = JOIN(arr, separator)
    
    ' Return the joined string of cell values in the specified order.
    JOIN_SORT = result
End Function

' Function: JOIN_NON_EMPTY
' Description: Joins the non-empty values of a cell range using a specified separator (optional).
' Parameters:
'   SourceRange: The range of cells to join non-empty values from.
'   Optional Separator: The separator character used to join the values (default is an empty string).
' Returns:
'   The joined string of non-empty cell values.
Function JOIN_NON_EMPTY(SourceRange As Range, Optional separator As String = "") As String
    Dim cell As Range
    Dim result As String
    
    result = ""
    
    ' Loop through each cell in the SourceRange.
    For Each cell In SourceRange
        ' Check if the cell is not empty before appending the value to the result string.
        If cell.value <> "" Then
            result = result & cell.value & separator
        End If
    Next cell
    
    ' Remove the trailing separator if it exists.
    If Len(separator) > 0 Then
        result = Left(result, Len(result) - Len(separator))
    End If
    
    ' Return the joined string of non-empty cell values.
    JOIN_NON_EMPTY = result
End Function

' Function: JOIN_NON_EMPTY_SORT
' Description: Joins the non-empty values of a cell range using a specified separator (optional) and in ascending (default) or descending order.
' Parameters:
'   SourceRange: The range of cells to join non-empty values from.
'   Optional Separator: The separator character used to join the values (default is an empty string).
'   Optional SortOrder: The sorting order of the joined values ("ASC" for ascending, "DESC" for descending).
' Returns:
'   The joined string of non-empty cell values in the specified order.
Function JOIN_NON_EMPTY_SORT(SourceRange As Range, Optional separator As String = "", Optional SortOrder As String = "ASC") As String
    Dim cell As Range
    Dim result As String
    Dim arr() As String
    
    result = ""
    
    ' Loop through each cell in the SourceRange and store the non-empty values in an array.
    For Each cell In SourceRange
        If cell.value <> "" Then
            ReDim Preserve arr(0 To UBound(arr) + 1)
            arr(UBound(arr)) = cell.value
        End If
    Next cell
    
    ' Sort the array based on the specified SortOrder.
    If UCase(SortOrder) = "DESC" Then
        SortArrayDescending arr
    Else
        SortArrayAscending arr
    End If
    
    ' Join the sorted array with the separator.
    result = JOIN(arr, separator)
    
    ' Return the joined string of non-empty cell values in the specified order.
    JOIN_NON_EMPTY_SORT = result
End Function

' Function: JOIN_UNIQUE
' Description: Joins the unique (non-duplicate and non-empty) values of a cell range using a specified separator (optional).
' Parameters:
'   SourceRange: The range of cells to join unique values from.
'   Optional Separator: The separator character used to join the values (default is an empty string).
' Returns:
'   The joined string of unique cell values.
Function JOIN_UNIQUE(SourceRange As Range, Optional separator As String = "") As String
    Dim cell As Range
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim result As String
    
    result = ""
    
    ' Loop through each cell in the SourceRange.
    For Each cell In SourceRange
        ' Ignore empty cells and add non-duplicate values to the dictionary.
        If cell.value <> "" And Not dict.Exists(cell.value) Then
            dict.Add cell.value, 1
            result = result & cell.value & separator
        End If
    Next cell
    
    ' Remove the trailing separator if it exists.
    If Len(separator) > 0 Then
        result = Left(result, Len(result) - Len(separator))
    End If
    
    ' Return the joined string of unique cell values.
    JOIN_UNIQUE = result
End Function

' Function: JOIN_UNIQUE_SORT
' Description: Joins the unique (non-duplicate and non-empty) values of a cell range using a specified separator (optional) and in ascending (default) or descending order.
' Parameters:
'   SourceRange: The range of cells to join unique values from.
'   Optional Separator: The separator character used to join the values (default is an empty string).
'   Optional SortOrder: The sorting order of the joined values ("ASC" for ascending, "DESC" for descending).
' Returns:
'   The joined string of unique cell values in the specified order.
Function JOIN_UNIQUE_SORT(SourceRange As Range, Optional separator As String = "", Optional SortOrder As String = "ASC") As String
    Dim cell As Range
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim result As String
    Dim arr() As String
    
    result = ""
    
    ' Loop through each cell in the SourceRange and store the unique values in an array.
    For Each cell In SourceRange
        If cell.value <> "" And Not dict.Exists(cell.value) Then
            dict.Add cell.value, 1
            ReDim Preserve arr(0 To UBound(arr) + 1)
            arr(UBound(arr)) = cell.value
        End If
    Next cell
    
    ' Sort the array based on the specified SortOrder.
    If UCase(SortOrder) = "DESC" Then
        SortArrayDescending arr
    Else
        SortArrayAscending arr
    End If
    
    ' Join the sorted array with the separator.
    result = JOIN(arr, separator)
    
    ' Return the joined string of unique cell values in the specified order.
    JOIN_UNIQUE_SORT = result
End Function

' Helper Function: SortArrayAscending
' Description: Sorts a string array in ascending order.
Private Function SortArrayAscending(ByRef arr() As String)
    Dim i As Long, j As Long
    Dim temp As String
    
    ' Bubble sort the array.
    For i = LBound(arr) To UBound(arr)
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Function

' Helper Function: SortArrayDescending
' Description: Sorts a string array in descending order.
Private Function SortArrayDescending(ByRef arr() As String)
    Dim i As Long, j As Long
    Dim temp As String
    
    ' Bubble sort the array.
    For i = LBound(arr) To UBound(arr)
        For j = i + 1 To UBound(arr)
            If arr(i) < arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Function

