Attribute VB_Name = "Module6"
' Function: EXTRACT_WORD
' Description: Returns the nth word from a string, with the option to specify up to 3 separators.
' Parameters:
'   TextString: The input string to extract the word from.
'   NthWord: The position of the word to be extracted.
'   Optional Separator1: The first separator character (default is a space).
'   Optional Separator2: The second separator character (default is an empty string).
'   Optional Separator3: The third separator character (default is an empty string).
' Returns:
'   The nth word from the input string.
Function EXTRACT_WORD(TextString As String, NthWord As Long, _
                      Optional Separator1 As String = " ", _
                      Optional Separator2 As String = "", _
                      Optional Separator3 As String = "") As String
    Dim arr() As String
    Dim result As String
    Dim separators As String
    
    ' Concatenate the provided separators.
    separators = Separator1 & Separator2 & Separator3
    
    ' Split the input string into an array based on the specified separators.
    arr = Split(TextString, separators)
    
    ' Check if the nth word is within the array bounds.
    If NthWord >= 1 And NthWord <= UBound(arr) + 1 Then
        ' Return the nth word from the array.
        EXTRACT_WORD = arr(NthWord - 1)
    Else
        ' If the nth word is out of range, return an empty string.
        EXTRACT_WORD = ""
    End If
End Function

' Function: COUNT_TEXT
' Description: Counts the number of occurrences of a value in a text string.
' Parameters:
'   TextString: The input text string to search for occurrences.
'   SearchValue: The value to count occurrences of.
' Returns:
'   The count of occurrences of the search value in the text string.
Function COUNT_TEXT(TextString As String, searchValue As String) As Long
    Dim count As Long
    Dim startIndex As Long
    Dim pos As Long
    
    count = 0
    startIndex = 1
    
    ' Loop through the text string to find all occurrences of the search value.
    Do While startIndex <= Len(TextString)
        pos = InStr(startIndex, TextString, searchValue, vbTextCompare)
        If pos = 0 Then
            ' If the search value is not found, exit the loop.
            Exit Do
        Else
            ' Increment the count and update the start index for the next search.
            count = count + 1
            startIndex = pos + Len(searchValue)
        End If
    Loop
    
    ' Return the count of occurrences.
    COUNT_TEXT = count
End Function


