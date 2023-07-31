Attribute VB_Name = "Module9"
' Function: RANDOM_INTEGER
' Description: Returns a random integer between two specified values (similar to RAND.BETWEEN but usable with older versions of Excel).
' Parameters:
'   MinValue: The minimum value of the random integer range.
'   MaxValue: The maximum value of the random integer range.
' Returns:
'   A random integer between MinValue and MaxValue (inclusive).
Function RANDOM_INTEGER(minValue As Long, maxValue As Long) As Long
    ' Check if the MinValue is greater than the MaxValue, and swap them if necessary.
    If minValue > maxValue Then
        Dim temp As Long
        temp = minValue
        minValue = maxValue
        maxValue = temp
    End If
    
    ' Calculate the random number within the specified range.
    RANDOM_INTEGER = Int((maxValue - minValue + 1) * Rnd) + minValue
End Function

' Function: NON_VOLATILE_RANDOM_INTEGER
' Description: Returns a non-volatile random integer between two specified values (the value is not volatile and does not change when you modify other cells).
' Parameters:
'   MinValue: The minimum value of the random integer range.
'   MaxValue: The maximum value of the random integer range.
' Returns:
'   A random integer between MinValue and MaxValue (inclusive).
Function NON_VOLATILE_RANDOM_INTEGER(minValue As Long, maxValue As Long) As Long
    ' Check if the MinValue is greater than the MaxValue, and swap them if necessary.
    If minValue > maxValue Then
        Dim temp As Long
        temp = minValue
        minValue = maxValue
        maxValue = temp
    End If
    
    ' Generate a unique seed value based on the current worksheet name and cell address.
    Dim seed As Long
    seed = CLng(Application.ThisWorkbook.Worksheets(Application.Caller.Parent.Name).Cells(Application.Caller.row, Application.Caller.Column).Address)
    
    ' Set the random seed value to make the random number non-volatile.
    Randomize seed
    
    ' Calculate the random number within the specified range.
    NON_VOLATILE_RANDOM_INTEGER = Int((maxValue - minValue + 1) * Rnd) + minValue
End Function

