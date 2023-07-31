Attribute VB_Name = "Module7"
' Function: CONVERT_TO_HOURS
' Description: Converts a time value in "hh:mm:ss" or "h:mm" format to a decimal number of hours.
' Parameters:
'   TimeString: The time value in "hh:mm:ss" or "h:mm" format.
' Returns:
'   The time value converted to a decimal number of hours.
Function CONVERT_TO_HOURS(TimeString As String) As Double
    Dim timeValue As Date
    
    ' Use the built-in TimeValue function to convert the string to a time value.
    timeValue = timeValue(TimeString)
    
    ' Convert the time value to a decimal number of hours.
    CONVERT_TO_HOURS = timeValue * 24
End Function

' Function: DATE_DIFF
' Description: Returns the difference in days between two dates.
' Parameters:
'   Date1: The first date.
'   Date2: The second date.
' Returns:
'   The difference in days between the two dates.
Function DATE_DIFF(Date1 As Date, Date2 As Date) As Long
    ' Use the built-in DateDiff function to calculate the difference in days.
    DATE_DIFF = DateDiff("d", Date1, Date2)
End Function

' Function: DATE_EU
' Description: Returns TRUE if Excel uses the EU date format (dd/mm), or FALSE if it uses the US date format (mm/dd).
' Parameters: None
' Returns:
'   TRUE if Excel uses EU date format, or FALSE if it uses US date format.
Function DATE_EU() As Boolean
    ' Check the date format setting in Excel.
    ' xlMDY indicates US date format, xlDMY indicates EU date format.
    DATE_EU = (Application.International(xlDateOrder) = xlDMY)
End Function

' Function: DATE_INVERT_DM
' Description: Returns a date with the day and month inverted (if possible), or the same date if inversion is not possible.
' Parameters:
'   InputDate: The input date to be inverted.
' Returns:
'   The date with the day and month inverted (if possible), or the same date if inversion is not possible.
Function DATE_INVERT_DM(inputDate As Date) As Date
    Dim dayValue As Integer
    Dim monthValue As Integer
    
    ' Extract the day and month components from the input date.
    dayValue = Day(inputDate)
    monthValue = Month(inputDate)
    
    ' Check if the inversion is possible (day value <= 12 and month value <= 12).
    If dayValue <= 12 And monthValue <= 12 Then
        ' Invert the day and month components and create the inverted date.
        DATE_INVERT_DM = DateSerial(Year(inputDate), monthValue, dayValue)
    Else
        ' If inversion is not possible, return the same date.
        DATE_INVERT_DM = inputDate
    End If
End Function

' Function: NB_DAYS_MONTH
' Description: Returns the number of days in a month based on a date.
' Parameters:
'   InputDate: The input date from which to determine the month.
' Returns:
'   The number of days in the month of the input date.
Function NB_DAYS_MONTH(inputDate As Date) As Integer
    ' Use the DateSerial function to get the first day of the next month and then subtract 1 to get the last day of the current month.
    NB_DAYS_MONTH = Day(DateSerial(Year(inputDate), Month(inputDate) + 1, 1) - 1)
End Function

' Function: ISO_WEEK_NUMBER
' Description: Returns the ISO week number based on a date (from 1900 to 2200).
' Parameters:
'   InputDate: The input date for which to determine the ISO week number.
' Returns:
'   The ISO week number for the input date.
Function ISO_WEEK_NUMBER(inputDate As Date) As Integer
    ' Use the DatePart function to get the ISO week number (ww).
    ISO_WEEK_NUMBER = DatePart("ww", inputDate, vbMonday, vbFirstFourDays)
End Function

' Function: ASCENSION_DATE
' Description: Returns the date of Ascension based on a year (or the year of a given date), from 1900 to 2200.
' Parameters:
'   Optional YearOrDate: The year for which to determine the date of Ascension, or a specific date from which to extract the year.
' Returns:
'   The date of Ascension for the specified year, or the year of the input date.
Function ASCENSION_DATE(Optional YearOrDate As Variant) As Date
    Dim inputYear As Integer
    
    ' If the YearOrDate parameter is provided, extract the year from it.
    If Not IsMissing(YearOrDate) Then
        If IsDate(YearOrDate) Then
            inputYear = Year(YearOrDate)
        Else
            inputYear = CInt(YearOrDate)
        End If
    Else
        ' If the YearOrDate parameter is not provided, use the current year.
        inputYear = Year(Date)
    End If
    
    ' Calculate the date of Ascension based on the input year.
    ASCENSION_DATE = DateSerial(inputYear, 1, 1) + 39 + (7 - Weekday(DateSerial(inputYear, 1, 1) + 39, vbMonday))
End Function

' Function: PENTECOST_MONDAY_DATE
' Description: Returns the date of Pentecost Monday based on a year (or the year of a given date), from 1900 to 2200.
' Parameters:
'   Optional YearOrDate: The year for which to determine the date of Pentecost Monday, or a specific date from which to extract the year.
' Returns:
'   The date of Pentecost Monday for the specified year, or the year of the input date.
Function PENTECOST_MONDAY_DATE(Optional YearOrDate As Variant) As Date
    Dim inputYear As Integer
    
    ' If the YearOrDate parameter is provided, extract the year from it.
    If Not IsMissing(YearOrDate) Then
        If IsDate(YearOrDate) Then
            inputYear = Year(YearOrDate)
        Else
            inputYear = CInt(YearOrDate)
        End If
    Else
        ' If the YearOrDate parameter is not provided, use the current year.
        inputYear = Year(Date)
    End If
    
    ' Calculate the date of Pentecost Monday based on the input year.
    PENTECOST_MONDAY_DATE = ASCENSION_DATE(inputYear) + 10
End Function

' Function: EASTER_DATE
' Description: Returns the date of Easter based on a year (or the year of a given date), from 1900 to 2200.
' Parameters:
'   Optional YearOrDate: The year for which to determine the date of Easter, or a specific date from which to extract the year.
' Returns:
'   The date of Easter for the specified year, or the year of the input date.
Function EASTER_DATE(Optional YearOrDate As Variant) As Date
    Dim inputYear As Integer
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
    
    ' If the YearOrDate parameter is provided, extract the year from it.
    If Not IsMissing(YearOrDate) Then
        If IsDate(YearOrDate) Then
            inputYear = Year(YearOrDate)
        Else
            inputYear = CInt(YearOrDate)
        End If
    Else
        ' If the YearOrDate parameter is not provided, use the current year.
        inputYear = Year(Date)
    End If
    
    ' Calculate the date of Easter based on the input year (Gauss's algorithm).
    a = inputYear Mod 19
    b = inputYear \ 100
    c = inputYear Mod 100
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
    
    EASTER_DATE = DateSerial(inputYear, (h + l - 7 * m + 114) \ 31, (p + 1) Mod 31 + 1)
End Function

' Function: IS_EASTER
' Description: Returns TRUE if the date (from 1900 to 2200) corresponds to Easter, or FALSE otherwise.
' Parameters:
'   TestDate: The date to check for Easter.
' Returns:
'   TRUE if the date corresponds to Easter, or FALSE otherwise.
Function IS_EASTER(TestDate As Date) As Boolean
    ' Use the EasterDate function to get the date of Easter for the year of the TestDate.
    ' If the TestDate matches the date of Easter, return TRUE; otherwise, return FALSE.
    IS_EASTER = (TestDate = EASTER_DATE(Year(TestDate)))
End Function

