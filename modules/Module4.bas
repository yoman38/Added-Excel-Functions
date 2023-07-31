Attribute VB_Name = "Module4"
' Function: DUPLICATES
' Description: Returns TRUE if the cell range contains duplicates (ignoring empty cells), or FALSE if there are no duplicates.
' Parameters:
'   SearchRange: The range of cells to check for duplicates.
' Returns:
'   True if there are duplicates, False otherwise.
Function DUPLICATES(searchRange As Range) As Boolean
    Dim cell As Range
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Loop through each cell in the SearchRange.
    For Each cell In searchRange
        ' Ignore empty cells.
        If cell.value <> "" Then
            ' If the value already exists in the dictionary, it's a duplicate.
            If dict.Exists(cell.value) Then
                DUPLICATES = True
                Exit Function
            Else
                dict.Add cell.value, 1
            End If
        End If
    Next cell
    
    ' If no duplicates found, return False.
    DUPLICATES = False
End Function

' Function: DUPLICATES_ADDRESSES
' Description: Returns the addresses of duplicates in a cell range (ignoring empty cells), with the option to specify the address separator (optional) and the value to return if no duplicates are found (optional).
' Parameters:
'   SearchRange: The range of cells to check for duplicates.
'   Optional Separator: The separator character used to join the addresses (default is an empty string).
'   Optional NoDuplicateValue: The value to return if no duplicates are found (default is an empty string).
' Returns:
'   The addresses of the duplicates joined by the separator or the NoDuplicateValue if no duplicates found.
Function DUPLICATES_ADDRESSES(searchRange As Range, Optional separator As String = "", Optional NoDuplicateValue As Variant = "") As String
    Dim cell As Range
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim addresses As String
    addresses = ""
    
    ' Loop through each cell in the SearchRange.
    For Each cell In searchRange
        ' Ignore empty cells.
        If cell.value <> "" Then
            ' If the value already exists in the dictionary, it's a duplicate.
            If dict.Exists(cell.value) Then
                addresses = addresses & cell.Address(False, False) & separator
            Else
                dict.Add cell.value, 1
            End If
        End If
    Next cell
    
    ' Remove the trailing separator if it exists.
    If Len(separator) > 0 Then
        addresses = Left(addresses, Len(addresses) - Len(separator))
    End If
    
    ' Return the addresses of the duplicates or the NoDuplicateValue if no duplicates found.
    If addresses = "" Then
        DUPLICATES_ADDRESSES = NoDuplicateValue
    Else
        DUPLICATES_ADDRESSES = addresses
    End If
End Function

' Function: DUPLICATES_LIST
' Description: Returns a list of duplicates in a cell range (ignoring empty cells), with the option to specify the list separator (optional) and the value to return if no duplicates are found (optional).
' Parameters:
'   SearchRange: The range of cells to check for duplicates.
'   Optional Separator: The separator character used to join the list (default is an empty string).
'   Optional NoDuplicateValue: The value to return if no duplicates are found (default is an empty string).
' Returns:
'   The list of duplicates joined by the separator or the NoDuplicateValue if no duplicates found.
Function DUPLICATES_LIST(searchRange As Range, Optional separator As String = "", Optional NoDuplicateValue As Variant = "") As String
    Dim cell As Range
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim duplicatesList As String
    duplicatesList = ""
    
    ' Loop through each cell in the SearchRange.
    For Each cell In searchRange
        ' Ignore empty cells.
        If cell.value <> "" Then
            ' If the value already exists in the dictionary, it's a duplicate.
            If dict.Exists(cell.value) Then
                ' Check if the value is already added to the duplicates list.
                If InStr(duplicatesList, cell.value & separator) = 0 Then
                    duplicatesList = duplicatesList & cell.value & separator
                End If
            Else
                dict.Add cell.value, 1
            End If
        End If
    Next cell
    
    ' Remove the trailing separator if it exists.
    If Len(separator) > 0 Then
        duplicatesList = Left(duplicatesList, Len(duplicatesList) - Len(separator))
    End If
    
    ' Return the list of duplicates or the NoDuplicateValue if no duplicates found.
    If duplicatesList = "" Then
        DUPLICATES_LIST = NoDuplicateValue
    Else
        DUPLICATES_LIST = duplicatesList
    End If
End Function

' Function: UNIQUE_LIST
' Description: Returns a list of unique values from a cell range (ignoring empty cells), with the option to specify the list separator (optional) and the sorting order (optional).
' Parameters:
'   SearchRange: The range of cells to get unique values.
'   Optional Separator: The separator character used to join the list (default is an empty string).
'   Optional SortOrder: The sorting order of the list (default is ascending order).
' Returns:
'   The list of unique values joined by the separator.
Function UNIQUE_LIST(searchRange As Range, Optional separator As String = "", Optional SortOrder As String = "ASC") As String
    Dim cell As Range
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim uniqueList As String
    uniqueList = ""
    
    ' Loop through each cell in the SearchRange.
    For Each cell In searchRange
        ' Ignore empty cells.
        If cell.value <> "" Then
            ' If the value is not already in the dictionary, add it as a unique value.
            If Not dict.Exists(cell.value) Then
                dict.Add cell.value, 1
                uniqueList = uniqueList & cell.value & separator
            End If
        End If
    Next cell
    
    ' Remove the trailing separator if it exists.
    If Len(separator) > 0 Then
        uniqueList = Left(uniqueList, Len(uniqueList) - Len(separator))
    End If
    
    ' Sort the list if requested.
    If UCase(SortOrder) = "ASC" Then
        uniqueList = SortStringAscending(uniqueList, separator)
    ElseIf UCase(SortOrder) = "DESC" Then
        uniqueList = SortStringDescending(uniqueList, separator)
    End If
    
    ' Return the list of unique values joined by the separator.
    UNIQUE_LIST = uniqueList
End Function

' Helper Function: SortStringAscending
' Description: Sorts a string in ascending order using a separator.
Private Function SortStringAscending(ByVal inputString As String, ByVal separator As String) As String
    Dim arr() As String
    arr = Split(inputString, separator)
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
    
    ' Rejoin the sorted array with the separator.
    SortStringAscending = JOIN(arr, separator)
End Function

' Helper Function: SortStringDescending
' Description: Sorts a string in descending order using a separator.
Private Function SortStringDescending(ByVal inputString As String, ByVal separator As String) As String
    Dim arr() As String
    arr = Split(inputString, separator)
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
    
    ' Rejoin the sorted array with the separator.
    SortStringDescending = JOIN(arr, separator)
End Function

