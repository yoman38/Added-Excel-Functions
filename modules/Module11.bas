Attribute VB_Name = "Module11"
' Function: regexExtract
' Description: Extracts one or more parts of a string using regular expressions.
' Parameters:
'   inputString: The input string from which to extract the parts.
'   regexPattern: The regular expression pattern to match and extract the parts.
'   Optional separator: The separator to join multiple extracted parts (default is a space).
' Returns:
'   A string containing the extracted parts joined by the separator.
Function regexExtract(inputString As String, regexPattern As String, Optional separator As String = " ") As String
    Dim regEx As Object
    Dim matches As Object
    Dim match As Object
    Dim extractedParts As String
    
    ' Create a regular expression object.
    Set regEx = CreateObject("VBScript.RegExp")
    
    ' Set the regular expression pattern.
    regEx.Pattern = regexPattern
    
    ' Test if the input string matches the regular expression pattern.
    If regEx.Test(inputString) Then
        ' Get all matches from the input string.
        Set matches = regEx.Execute(inputString)
        
        ' Loop through each match and concatenate the extracted parts with the separator.
        For Each match In matches
            extractedParts = extractedParts & match.value & separator
        Next match
        
        ' Remove the trailing separator and return the extracted parts.
        regexExtract = Left(extractedParts, Len(extractedParts) - Len(separator))
    Else
        ' Return an empty string if there are no matches.
        regexExtract = ""
    End If
End Function

' Function: regexMatch
' Description: Tests if a string matches a regular expression pattern.
' Parameters:
'   inputString: The input string to be tested against the regular expression pattern.
'   regexPattern: The regular expression pattern to match.
' Returns:
'   True if the input string matches the regular expression pattern, False otherwise.
Function regexMatch(inputString As String, regexPattern As String) As Boolean
    Dim regEx As Object
    
    ' Create a regular expression object.
    Set regEx = CreateObject("VBScript.RegExp")
    
    ' Set the regular expression pattern.
    regEx.Pattern = regexPattern
    
    ' Test if the input string matches the regular expression pattern.
    regexMatch = regEx.Test(inputString)
End Function

' Function: regexReplace
' Description: Replaces one or more parts of a string using regular expressions.
' Parameters:
'   inputString: The input string in which to replace the parts.
'   regexPattern: The regular expression pattern to match the parts to be replaced.
'   replacement: The string to replace the matched parts with.
' Returns:
'   The input string with the matched parts replaced with the specified replacement.
Function regexReplace(inputString As String, regexPattern As String, replacement As String) As String
    Dim regEx As Object
    
    ' Create a regular expression object.
    Set regEx = CreateObject("VBScript.RegExp")
    
    ' Set the regular expression pattern.
    regEx.Pattern = regexPattern
    
    ' Replace the matched parts in the input string with the specified replacement.
    regexReplace = regEx.Replace(inputString, replacement)
End Function

