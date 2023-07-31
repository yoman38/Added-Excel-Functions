Attribute VB_Name = "Module14"
' Function: isInt
' Description: Returns True if the value is an integer number or False otherwise.
' Parameters:
'   value: The value to check if it's an integer.
' Returns:
'   True if the value is an integer, False otherwise.
Function isInt(ByVal value As Variant) As Boolean
    If IsNumeric(value) Then
        isInt = (Int(value) = value)
    Else
        isInt = False
    End If
End Function

' Function: intRand
' Description: Returns a random integer number between two specified values.
' Parameters:
'   minValue: The minimum value of the random range (inclusive).
'   maxValue: The maximum value of the random range (inclusive).
' Returns:
'   The random integer between minValue and maxValue.
Function intRand(ByVal minValue As Integer, ByVal maxValue As Integer) As Integer
    Randomize
    intRand = Int((maxValue - minValue + 1) * Rnd + minValue)
End Function

' Function: cellsSearch
' Description: Searches for a value in a range of cells and returns an array of addresses
'              containing the searched value.
' Parameters:
'   searchValue: The value to search for in the range of cells.
'   searchRange: The range of cells to search in.
' Returns:
'   An array containing the addresses of all cells containing the searchValue.
Function cellsSearch(ByVal searchValue As Variant, ByVal searchRange As Range) As Variant
    Dim foundCells As Range
    Dim cell As Range
    Dim result() As Variant
    Dim i As Long
    
    On Error Resume Next
    Set foundCells = searchRange.SpecialCells(xlCellTypeConstants, xlTextValues).Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo 0
    
    If foundCells Is Nothing Then
        ReDim result(1 To 1, 1 To 1)
        result(1, 1) = "Not Found"
    Else
        ReDim result(1 To foundCells.Cells.count, 1 To 1)
        i = 1
        For Each cell In foundCells
            result(i, 1) = cell.Address
            i = i + 1
        Next cell
    End If
    
    cellsSearch = result
End Function

' Function: isoWeekNum
' Description: Returns the ISO week number based on a date (from 1900 to 2200).
' Parameters:
'   inputDate: The date for which to calculate the ISO week number.
' Returns:
'   The ISO week number.
Function isoWeekNum(ByVal inputDate As Date) As Integer
    Dim isoWeekNumber As Integer
    
    isoWeekNumber = WorksheetFunction.WeekNum(inputDate, 21)
    isoWeekNum = IIf(isoWeekNumber = 53 And Month(inputDate) = 1, 1, isoWeekNumber)
End Function

' Function: nbDaysMonth
' Description: Returns the number of days in a month based on a date.
' Parameters:
'   inputDate: The date for which to calculate the number of days in the month.
' Returns:
'   The number of days in the month.
Function nbDaysMonth(ByVal inputDate As Date) As Integer
    nbDaysMonth = Day(DateSerial(Year(inputDate), Month(inputDate) + 1, 0))
End Function

' Function: euDate
' Description: Returns True if Excel uses the EU date format (dd/mm), False if Excel uses the US date format (mm/dd).
' Parameters: None
' Returns:
'   True if Excel uses the EU date format, False if Excel uses the US date format.
Function euDate() As Boolean
    euDate = (Application.International(xlDateOrder) = 1)
End Function

' Function: easterDate
' Description: Returns the date of Easter based on a year (or the year of a given date, from 1900 to 2200).
' Parameters:
'   Optional inputYear As Variant: The year for which to calculate the date of Easter. If not provided,
'                                  the year of the current date is used.
' Returns:
'   The date of Easter.
Function easterDate(Optional ByVal inputYear As Variant) As Date
    Dim yearNumber As Integer
    Dim a As Integer
    Dim b As Integer
    Dim c As Integer
    Dim d As Integer
    Dim e As Integer
    Dim f As Integer
    Dim g As Integer
    Dim h As Integer
    Dim i As Integer
    Dim k As Integer
    Dim l As Integer
    Dim m As Integer
    Dim p As Integer
    Dim easterMonth As Integer
    Dim easterDay As Integer
    
    If IsMissing(inputYear) Then
        yearNumber = Year(Date)
    Else
        yearNumber = inputYear
    End If
    
    a = yearNumber Mod 19
    b = yearNumber \ 100
    c = yearNumber Mod 100
    d = b \ 4
    e = b Mod 4
    f = (b + 8) \ 25
    g = (b - f + 1) \ 3
    h = (19 * a + b - d - g + 15) Mod 30
    i = c \ 4
    k = c Mod 4
    l = (32 + 2 * e + 2 * i - h - k) Mod 7
    m = (a + 11 * h + 22 * l) \ 451
    p = (h + l - 7 * m + 114) Mod 31
    
    easterMonth = (h + l - 7 * m + 114) \ 31
    easterDay = p + 1
    
    easterDate = DateSerial(yearNumber, easterMonth, easterDay)
End Function

