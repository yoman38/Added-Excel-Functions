Attribute VB_Name = "Module1"
Function MAX_IF_COLOR(searchRange As Range, RefCell As Range) As Double
    Dim cell As Range
    Dim maxVal As Double
    maxVal = -1E+307
    For Each cell In searchRange
        If cell.Interior.Color = RefCell.Interior.Color And cell.value > maxVal Then
            maxVal = cell.value
        End If
    Next cell
    MAX_IF_COLOR = maxVal
End Function

Function MIN_IF_COLOR(searchRange As Range, RefCell As Range) As Double
    Dim cell As Range
    Dim minVal As Double
    minVal = 1E+307
    For Each cell In searchRange
        If cell.Interior.Color = RefCell.Interior.Color And cell.value < minVal Then
            minVal = cell.value
        End If
    Next cell
    MIN_IF_COLOR = minVal
End Function

Function AVERAGE_IF_COLOR(searchRange As Range, RefCell As Range) As Double
    Dim cell As Range
    Dim sumVal As Double
    Dim count As Integer
    sumVal = 0
    count = 0
    For Each cell In searchRange
        If cell.Interior.Color = RefCell.Interior.Color And IsNumeric(cell.value) Then
            sumVal = sumVal + cell.value
            count = count + 1
        End If
    Next cell
    If count > 0 Then
        AVERAGE_IF_COLOR = sumVal / count
    Else
        AVERAGE_IF_COLOR = "No cells with specified color found"
    End If
End Function

Function COUNT_COLORED(searchRange As Range) As Integer
    Dim cell As Range
    Dim count As Integer
    count = 0
    For Each cell In searchRange
        If cell.Interior.Color <> RGB(255, 255, 255) And cell.Interior.Color <> 0 Then
            count = count + 1
        End If
    Next cell
    COUNT_COLORED = count
End Function

Function COUNT_IF_COLOR(searchRange As Range, RefCell As Range) As Integer
    Dim cell As Range
    Dim count As Integer
    count = 0
    For Each cell In searchRange
        If cell.Interior.Color = RefCell.Interior.Color Then
            count = count + 1
        End If
    Next cell
    COUNT_IF_COLOR = count
End Function
Function SUM_IF_COLORED(RangeData As Range) As Double
    Dim cell As Range
    Dim Sum As Double
    Sum = 0
    For Each cell In RangeData
        If cell.Interior.ColorIndex <> xlNone And cell.Interior.Color <> RGB(255, 255, 255) And IsNumeric(cell.value) Then
            Sum = Sum + cell.value
        End If
    Next cell
    SUM_IF_COLORED = Sum
End Function
Function SUM_IF_COLOR(RangeData As Range, CriteriaRange As Range) As Double
    Dim cell As Range
    Dim Sum As Double
    Sum = 0
    For Each cell In RangeData
        If cell.Interior.Color = CriteriaRange.Interior.Color And IsNumeric(cell.value) Then
            Sum = Sum + cell.value
        End If
    Next cell
    SUM_IF_COLOR = Sum
End Function
Function COLOR_NUMBER(CellRef As Range) As Long
    COLOR_NUMBER = CellRef.Interior.Color
End Function
Function HEX_TO_COLOR_NUMBER(HexColor As String) As Long
    On Error GoTo ErrorHandler
    HEX_TO_COLOR_NUMBER = RGB(Application.Hex2Dec(Mid(HexColor, 2, 2)), _
                              Application.Hex2Dec(Mid(HexColor, 4, 2)), _
                              Application.Hex2Dec(Mid(HexColor, 6, 2)))
    Exit Function
ErrorHandler:
    HEX_TO_COLOR_NUMBER = -1
End Function

