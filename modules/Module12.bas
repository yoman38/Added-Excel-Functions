Attribute VB_Name = "Module12"
' Function: colLetter
' Description: Converts a column number to its corresponding letter representation.
' Parameters:
'   columnNumber: The column number to convert to letter(s).
' Returns:
'   The corresponding letter(s) representing the column number.
Function colLetter(columnNumber As Long) As String
    Dim dividend As Long
    Dim moduloResult As Long
    Dim columnName As String
    
    If columnNumber <= 0 Then
        colLetter = ""
        Exit Function
    End If
    
    dividend = columnNumber
    Do
        moduloResult = (dividend - 1) Mod 26
        columnName = Chr(65 + moduloResult) & columnName
        dividend = (dividend - moduloResult) \ 26
    Loop While dividend > 0
    
    colLetter = columnName
End Function

' Function: colNum
' Description: Converts a column letter(s) to its corresponding number representation.
' Parameters:
'   columnLetter: The column letter(s) to convert to a column number.
' Returns:
'   The corresponding column number.
Function colNum(columnLetter As String) As Long
    Dim char As String
    Dim result As Long
    Dim position As Long
    Dim power As Long
    
    If Len(columnLetter) = 0 Then
        colNum = 0
        Exit Function
    End If
    
    result = 0
    power = 1
    For position = Len(columnLetter) To 1 Step -1
        char = UCase(Mid(columnLetter, position, 1))
        result = result + (Asc(char) - 64) * power
        power = power * 26
    Next position
    
    colNum = result
End Function

