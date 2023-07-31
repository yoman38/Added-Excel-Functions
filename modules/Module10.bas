Attribute VB_Name = "Module10"
' Function: arrayAdd
' Description: Increases the size of an array by 1 and adds a value to the last position.
' Parameters:
'   arr: The input array to which the value will be added.
'   value: The value to be added to the array.
' Returns:
'   The updated array with the new value added.
Function arrayAdd(arr() As Variant, value As Variant) As Variant()
    Dim newArray() As Variant
    Dim i As Long
    
    ' Increase the size of the array by 1.
    ReDim newArray(1 To UBound(arr) + 1)
    
    ' Copy the existing elements from the original array to the new array.
    For i = 1 To UBound(arr)
        newArray(i) = arr(i)
    Next i
    
    ' Add the new value to the last position in the new array.
    newArray(UBound(newArray)) = value
    
    ' Return the updated array with the new value added.
    arrayAdd = newArray
End Function

' Function: arrayCount
' Description: Returns the number of times the searched value is present in the array.
' Parameters:
'   arr: The input array to search for the value.
'   searchValue: The value to be counted in the array.
' Returns:
'   The count of occurrences of the searchValue in the array.
Function arrayCount(arr() As Variant, searchValue As Variant) As Long
    Dim count As Long
    Dim i As Long
    
    ' Initialize the count to 0.
    count = 0
    
    ' Loop through each element in the array and count the occurrences of the searchValue.
    For i = 1 To UBound(arr)
        If arr(i) = searchValue Then
            count = count + 1
        End If
    Next i
    
    ' Return the count of occurrences of the searchValue in the array.
    arrayCount = count
End Function

' Function: arrayDebug
' Description: Displays the content (or part of the content) of an array in a MsgBox.
' Parameters:
'   arr: The input array to be displayed in the MsgBox.
'   Optional startIdx: The starting index to display the array content (default is 1).
'   Optional endIdx: The ending index to display the array content (default is the last index of the array).
Sub arrayDebug(arr() As Variant, Optional startIdx As Long = 1, Optional endIdx As Long = -1)
    Dim result As String
    Dim i As Long
    
    ' Set the default value for endIdx to the last index of the array if not provided.
    If endIdx = -1 Then
        endIdx = UBound(arr)
    End If
    
    ' Build the string representation of the array content within the specified range.
    For i = startIdx To endIdx
        result = result & arr(i) & vbCrLf
    Next i
    
    ' Display the array content in a MsgBox.
    MsgBox result, vbInformation, "Array Debug"
End Sub

' Function: arrayDuplicates
' Description: Returns True if the array contains duplicates, or False if it contains no duplicates.
' Parameters:
'   arr: The input array to check for duplicates.
' Returns:
'   True if the array contains duplicates, or False if it contains no duplicates.
Function arrayDuplicates(arr() As Variant) As Boolean
    Dim values As Object
    Set values = CreateObject("Scripting.Dictionary")
    Dim i As Long
    
    ' Loop through each element in the array and add it to the dictionary.
    ' If the element already exists in the dictionary, it is a duplicate.
    For i = 1 To UBound(arr)
        If Not values.Exists(arr(i)) Then
            values(arr(i)) = True
        Else
            arrayDuplicates = True
            Exit Function
        End If
    Next i
    
    ' Return False if no duplicates were found in the array.
    arrayDuplicates = False
End Function

' Function: arrayDuplicatesDelete
' Description: Removes all duplicates from an array.
' Parameters:
'   arr: The input array from which duplicates will be removed.
' Returns:
'   The array with duplicates removed.
Function arrayDuplicatesDelete(arr() As Variant) As Variant()
    Dim values As Object
    Set values = CreateObject("Scripting.Dictionary")
    Dim i As Long
    Dim newArray() As Variant
    Dim newIndex As Long
    
    ' Loop through each element in the array and add it to the dictionary.
    ' If the element already exists in the dictionary, it is a duplicate and will be skipped.
    For i = 1 To UBound(arr)
        If Not values.Exists(arr(i)) Then
            values(arr(i)) = True
        End If
    Next i
    
    ' Initialize the new array with the size of the dictionary (non-duplicate elements).
    ReDim newArray(1 To values.count)
    
    ' Loop through the dictionary keys and populate the new array with non-duplicate elements.
    newIndex = 1
    For i = 0 To values.count - 1
        newIndex = newIndex + 1
        newArray(newIndex) = values.Keys(i)
    Next i
    
    ' Return the array with duplicates removed.
    arrayDuplicatesDelete = newArray
End Function

' Function: arrayDuplicatesList
' Description: Counts the number of times each value is present in the array and adds a second dimension to the array to record these values (1 = unique, 2 = double value, etc.).
' Parameters:
'   arr: The input array to count the occurrences of each value.
' Returns:
'   The array with a second dimension indicating the count of occurrences for each value.
Function arrayDuplicatesList(arr() As Variant) As Variant()
    Dim values As Object
    Set values = CreateObject("Scripting.Dictionary")
    Dim i As Long
    
    ' Loop through each element in the array and add it to the dictionary.
    ' If the element already exists in the dictionary, increase its count by 1.
    For i = 1 To UBound(arr)
        If Not values.Exists(arr(i)) Then
            values(arr(i)) = 1
        Else
            values(arr(i)) = values(arr(i)) + 1
        End If
    Next i
    
    ' Initialize the new array with the size of the dictionary (number of unique values).
    ReDim newArray(1 To values.count, 1 To 2)
    
    ' Loop through the dictionary keys and populate the new array with the values and their counts.
    For i = 0 To values.count - 1
        newArray(i + 1, 1) = values.Keys(i)
        newArray(i + 1, 2) = values.Items(i)
    Next i
    
    ' Return the array with a second dimension indicating the count of occurrences for each value.
    arrayDuplicatesList = newArray
End Function

' Function: arrayEmpty
' Description: Returns True if the array is empty, or False if it is not empty.
' Parameters:
'   arr: The input array to check for emptiness.
' Returns:
'   True if the array is empty, or False if it is not empty.
Function arrayEmpty(arr() As Variant) As Boolean
    ' Return True if the array has no elements (is empty).
    arrayEmpty = UBound(arr) < LBound(arr)
End Function

' Function: arrayPos
' Description: Returns the (first) position of the searched value in the array or returns -1 if the value was not found.
' Parameters:
'   arr: The input array to search for the value.
'   searchValue: The value to be searched in the array.
' Returns:
'   The position of the searchValue in the array, or -1 if the value was not found.
Function arrayPos(arr() As Variant, searchValue As Variant) As Long
    Dim i As Long
    
    ' Loop through each element in the array and find the first occurrence of the searchValue.
    For i = 1 To UBound(arr)
        If arr(i) = searchValue Then
            ' Return the position of the searchValue in the array.
            arrayPos = i
            Exit Function
        End If
    Next i
    
    ' Return -1 if the searchValue was not found in the array.
    arrayPos = -1
End Function

' Function: arrayRandomize
' Description: Randomly shuffles the values of an array.
' Parameters:
'   arr: The input array to be shuffled.
' Returns:
'   The array with values randomly shuffled.
Function arrayRandomize(arr() As Variant) As Variant()
    Dim temp As Variant
    Dim randomIndex As Long
    Dim i As Long
    
    ' Use the Rnd function to generate a random seed value for shuffling.
    Randomize
    
    ' Loop through each element in the array and swap it with a random element.
    ' This process randomizes the order of the elements.
    For i = 1 To UBound(arr)
        randomIndex = Int((UBound(arr) - LBound(arr) + 1) * Rnd + LBound(arr))
        temp = arr(randomIndex)
        arr(randomIndex) = arr(i)
        arr(i) = temp
    Next i
    
    ' Return the array with values randomly shuffled.
    arrayRandomize = arr
End Function

' Function: arraySortAsc
' Description: Sorts the values of an array in ascending order.
' Parameters:
'   arr: The input array to be sorted.
' Returns:
'   The array with values sorted in ascending order.
Function arraySortAsc(arr() As Variant) As Variant()
    ' Use the built-in VBA function to sort the array in ascending order.
    VBA.Sort arr
    
    ' Return the array with values sorted in ascending order.
    arraySortAsc = arr
End Function

' Function: arraySortDesc
' Description: Sorts the values of an array in descending order.
' Parameters:
'   arr: The input array to be sorted.
' Returns:
'   The array with values sorted in descending order.
Function arraySortDesc(arr() As Variant) As Variant()
    ' Use the built-in VBA function to sort the array in descending order.
    VBA.Sort arr, Descending:=True
    
    ' Return the array with values sorted in descending order.
    arraySortDesc = arr
End Function

' Function: arrayMax
' Description: Returns the largest numeric value present in the array.
' Parameters:
'   arr: The input array to find the maximum value.
' Returns:
'   The largest numeric value in the array.
Function arrayMax(arr() As Variant) As Variant
    Dim maxVal As Variant
    Dim i As Long
    
    ' Initialize the maxVal with the first element of the array.
    maxVal = arr(1)
    
    ' Loop through each element in the array and update maxVal if a larger value is found.
    For i = 2 To UBound(arr)
        If arr(i) > maxVal Then
            maxVal = arr(i)
        End If
    Next i
    
    ' Return the largest numeric value in the array.
    arrayMax = maxVal
End Function

' Function: arrayMin
' Description: Returns the smallest numeric value present in the array.
' Parameters:
'   arr: The input array to find the minimum value.
' Returns:
'   The smallest numeric value in the array.
Function arrayMin(arr() As Variant) As Variant
    Dim minVal As Variant
    Dim i As Long
    
    ' Initialize the minVal with the first element of the array.
    minVal = arr(1)
    
    ' Loop through each element in the array and update minVal if a smaller value is found.
    For i = 2 To UBound(arr)
        If arr(i) < minVal Then
            minVal = arr(i)
        End If
    Next i
    
    ' Return the smallest numeric value in the array.
    arrayMin = minVal
End Function

' Function: arrayNumDelete
' Description: Removes a value from an array based on its position in the array.
' Parameters:
'   arr: The input array from which the value will be removed.
'   position: The position of the value to be removed (1-based index).
' Returns:
'   The array with the specified value removed.
Function arrayNumDelete(arr() As Variant, position As Long) As Variant()
    Dim newArray() As Variant
    Dim i As Long
    Dim newIndex As Long
    
    ' Check if the position is within the valid range of the array.
    If position < 1 Or position > UBound(arr) Then
        ' Invalid position, return the original array unchanged.
        arrayNumDelete = arr
        Exit Function
    End If
    
    ' Initialize the new array with a size one less than the original array.
    ReDim newArray(1 To UBound(arr) - 1)
    
    ' Loop through the elements in the original array and populate the new array,
    ' excluding the element at the specified position.
    newIndex = 1
    For i = 1 To UBound(arr)
        If i <> position Then
            newArray(newIndex) = arr(i)
            newIndex = newIndex + 1
        End If
    Next i
    
    ' Return the array with the specified value removed.
    arrayNumDelete = newArray
End Function

' Function: arrayValuesDelete
' Description: Removes all occurrences of a value from an array.
' Parameters:
'   arr: The input array from which the values will be removed.
'   value: The value to be removed from the array.
' Returns:
'   The array with all occurrences of the specified value removed.
Function arrayValuesDelete(arr() As Variant, value As Variant) As Variant()
    Dim newArray() As Variant
    Dim i As Long
    Dim newIndex As Long
    
    ' Initialize the new array with a size equal to the original array.
    ReDim newArray(1 To UBound(arr))
    
    ' Loop through the elements in the original array and populate the new array,
    ' excluding any occurrences of the specified value.
    newIndex = 1
    For i = 1 To UBound(arr)
        If arr(i) <> value Then
            newArray(newIndex) = arr(i)
            newIndex = newIndex + 1
        End If
    Next i
    
    ' If the new array is not empty, resize it to remove any unused elements.
    If newIndex > 1 Then
        ReDim Preserve newArray(1 To newIndex - 1)
    End If
    
    ' Return the array with all occurrences of the specified value removed.
    arrayValuesDelete = newArray
End Function

' Function: inArray
' Description: Returns True if the value is found in the array, or False if it is not found.
' Parameters:
'   arr: The input array to search for the value.
'   value: The value to be searched in the array.
' Returns:
'   True if the value is found in the array, or False if it is not found.
Function inArray(arr() As Variant, value As Variant) As Boolean
    Dim i As Long
    
    ' Loop through each element in the array and check if the value matches any element.
    For i = 1 To UBound(arr)
        If arr(i) = value Then
            ' Return True if the value is found in the array.
            inArray = True
            Exit Function
        End If
    Next i
    
    ' Return False if the value is not found in the array.
    inArray = False
End Function

' Function: array2dDebug
' Description: Displays the content (or part of the content) of a 2D array in a MsgBox.
' Parameters:
'   arr: The input 2D array to be displayed in the MsgBox.
'   Optional startRow: The starting row index to display the array content (default is 1).
'   Optional endRow: The ending row index to display the array content (default is the last row of the array).
'   Optional startCol: The starting column index to display the array content (default is 1).
'   Optional endCol: The ending column index to display the array content (default is the last column of the array).
Sub array2dDebug(arr() As Variant, Optional startRow As Long = 1, Optional endRow As Long = -1, _
                 Optional startCol As Long = 1, Optional endCol As Long = -1)
    Dim result As String
    Dim i As Long, j As Long
    Dim rowCount As Long, colCount As Long
    Dim maxRows As Long, maxCols As Long
    
    ' Get the number of rows and columns in the array.
    rowCount = UBound(arr, 1) - LBound(arr, 1) + 1
    colCount = UBound(arr, 2) - LBound(arr, 2) + 1
    
    ' Set the default values for endRow and endCol if not provided.
    If endRow = -1 Then
        endRow = rowCount
    End If
    If endCol = -1 Then
        endCol = colCount
    End If
    
    ' Limit the values of startRow, endRow, startCol, and endCol to be within the array's dimensions.
    startRow = Application.WorksheetFunction.Max(1, startRow)
    endRow = Application.WorksheetFunction.Min(endRow, rowCount)
    startCol = Application.WorksheetFunction.Max(1, startCol)
    endCol = Application.WorksheetFunction.Min(endCol, colCount)
    
    ' Calculate the maximum number of rows and columns to display in the MsgBox.
    maxRows = Application.WorksheetFunction.Min(endRow - startRow + 1, 10) ' Display a maximum of 10 rows.
    maxCols = Application.WorksheetFunction.Min(endCol - startCol + 1, 5) ' Display a maximum of 5 columns.
    
    ' Build the string representation of the array content within the specified range.
    For i = startRow To startRow + maxRows - 1
        For j = startCol To startCol + maxCols - 1
            result = result & arr(i, j) & vbTab ' Separate columns with tabs.
        Next j
        result = result & vbCrLf ' Move to the next row.
    Next i
    
    ' Display the array content in a MsgBox.
    MsgBox result, vbInformation, "2D Array Debug"
End Sub


