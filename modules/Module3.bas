Attribute VB_Name = "Module3"
' Function: COUNT_IF_REGEX
' Description: Counts the number of cells that match a regular expression.
' Parameters:
'   SearchRange: The range of cells to search for matches.
'   Pattern: The regular expression pattern to match.
' Returns:
'   The number of cells that match the regular expression.
Function COUNT_IF_REGEX(searchRange As Range, Pattern As String) As Long
    Dim cell As Range
    Dim countMatches As Long
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    countMatches = 0
    
    ' Set the regular expression pattern.
    regEx.Pattern = Pattern
    
    ' Loop through each cell in the SearchRange.
    For Each cell In searchRange
        ' Check if the cell value matches the regular expression pattern.
        If regEx.Test(cell.value) Then
            countMatches = countMatches + 1
        End If
    Next cell
    
    ' Return the number of matching cells.
    COUNT_IF_REGEX = countMatches
End Function

' Function: REGEX_EXTRACT
' Description: Extracts one or more parts of a string using regular expressions, and allows specifying the separator (optional).
' Parameters:
'   InputString: The input string to extract parts from.
'   Pattern: The regular expression pattern to use for extraction.
'   Optional Separator: The separator character used to join the extracted parts (default is an empty string).
' Returns:
'   The extracted parts of the input string joined by the separator.
Function REGEX_EXTRACT(inputString As String, Pattern As String, Optional separator As String = "") As String
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    Dim matches As Object
    Dim result As String
    result = ""
    
    ' Set the regular expression pattern.
    regEx.Pattern = Pattern
    
    ' Execute the regular expression matching on the input string.
    Set matches = regEx.Execute(inputString)
    
    ' Concatenate the matched parts with the separator.
    Dim match As Object
    For Each match In matches
        result = result & match.value & separator
    Next match
    
    ' Remove the trailing separator if it exists.
    If Len(separator) > 0 Then
        result = Left(result, Len(result) - Len(separator))
    End If
    
    ' Return the extracted parts joined by the separator.
    REGEX_EXTRACT = result
End Function

' Function: REGEX_MATCH
' Description: Checks if a string matches a regular expression.
' Parameters:
'   InputString: The input string to check for a match.
'   Pattern: The regular expression pattern to use for matching.
' Returns:
'   True if the string matches the regular expression, False otherwise.
Function REGEX_MATCH(inputString As String, Pattern As String) As Boolean
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    ' Set the regular expression pattern.
    regEx.Pattern = Pattern
    
    ' Check if the input string matches the regular expression pattern.
    REGEX_MATCH = regEx.Test(inputString)
End Function

' Function: REGEX_REPLACE
' Description: Replaces one or more parts of a string using regular expressions.
' Parameters:
'   InputString: The input string in which replacements will be made.
'   Pattern: The regular expression pattern to use for replacements.
'   Replacement: The replacement string to use for matched parts.
' Returns:
'   The input string with matched parts replaced by the replacement string.
Function REGEX_REPLACE(inputString As String, Pattern As String, replacement As String) As String
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    ' Set the regular expression pattern.
    regEx.Pattern = Pattern
    
    ' Execute the regular expression replacement on the input string.
    REGEX_REPLACE = regEx.Replace(inputString, replacement)
End Function

' Function: SUM_IF_REGEX
' Description: Calculates the sum of cells that match a regular expression (ignoring non-numeric values).
' Parameters:
'   SearchRange: The range of cells to calculate the sum.
'   Pattern: The regular expression pattern to match.
' Returns:
'   The sum of the cells that match the regular expression (ignoring non-numeric values).
Function SUM_IF_REGEX(searchRange As Range, Pattern As String) As Double
    Dim cell As Range
    Dim sumVal As Double
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    sumVal = 0
    
    ' Set the regular expression pattern.
    regEx.Pattern = Pattern
    
    ' Loop through each cell in the SearchRange.
    For Each cell In searchRange
        ' Check if the cell value matches the regular expression pattern and if the cell's value is numeric.
        If regEx.Test(cell.value) And IsNumeric(cell.value) Then
            sumVal = sumVal + cell.value
        End If
    Next cell
    
    ' Return the sum of the cells that match the regular expression.
    SUM_IF_REGEX = sumVal
End Function

