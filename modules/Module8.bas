Attribute VB_Name = "Module8"
' Function: IS_EMAIL
' Description: Returns TRUE if the input string is a valid email address, or FALSE otherwise.
' Parameters:
'   EmailAddress: The input string to check for a valid email address.
' Returns:
'   TRUE if the input string is a valid email address, or FALSE otherwise.
Function IS_EMAIL(EmailAddress As String) As Boolean
    Dim regexPattern As String
    
    ' Define the regular expression pattern for email validation.
    regexPattern = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    
    ' Use the built-in Like operator to check if the input string matches the pattern.
    IS_EMAIL = (EmailAddress Like regexPattern)
End Function

' Function: IS_URL
' Description: Returns TRUE if the input string is a valid URL, or FALSE otherwise.
' Parameters:
'   URL: The input string to check for a valid URL.
' Returns:
'   TRUE if the input string is a valid URL, or FALSE otherwise.
Function IS_URL(url As String) As Boolean
    Dim regexPattern As String
    
    ' Define the regular expression pattern for URL validation.
    regexPattern = "^(https?|ftp)://[^\s/$.?#].[^\s]*$"
    
    ' Use the built-in Like operator to check if the input string matches the pattern.
    IS_URL = (url Like regexPattern)
End Function

' Function: HTML_TABLE
' Description: Assembles the values from a range of cells into a simple HTML table.
' Parameters:
'   SourceRange: The range of cells to include in the HTML table.
' Returns:
'   The HTML code representing the simple table.
Function HTML_TABLE(SourceRange As Range) As String
    Dim row As Range
    Dim cell As Range
    Dim result As String
    
    ' Initialize the result with the starting HTML table tag.
    result = "<table>"
    
    ' Loop through each row in the SourceRange.
    For Each row In SourceRange.Rows
        result = result & "<tr>"
        ' Loop through each cell in the row and add the cell value as a table data (td) element.
        For Each cell In row.Cells
            result = result & "<td>" & cell.value & "</td>"
        Next cell
        result = result & "</tr>"
    Next row
    
    ' Add the closing HTML table tag.
    result = result & "</table>"
    
    ' Return the assembled HTML table.
    HTML_TABLE = result
End Function

' Function: ADVANCED_HTML_TABLE
' Description: Assembles a range of cells into an HTML table while preserving the main formatting.
' Parameters:
'   SourceRange: The range of cells to include in the HTML table.
' Returns:
'   The HTML code representing the advanced table.
Function ADVANCED_HTML_TABLE(SourceRange As Range) As String
    Dim row As Range
    Dim cell As Range
    Dim result As String
    Dim rowColor As String
    
    ' Initialize the result with the starting HTML table tag and other attributes.
    result = "<table border='1' cellpadding='5' cellspacing='0'>"
    
    ' Loop through each row in the SourceRange.
    For Each row In SourceRange.Rows
        ' Set the row background color based on the cell interior color of the first cell in the row.
        rowColor = RGBToHex(row.Cells(1).Interior.Color)
        
        result = result & "<tr style='background-color: " & rowColor & "'>"
        
        ' Loop through each cell in the row and add the cell value as a table data (td) element with formatting.
        For Each cell In row.Cells
            result = result & "<td style='font-weight: bold; font-size: 12pt; color: white;'>" & cell.value & "</td>"
        Next cell
        
        result = result & "</tr>"
    Next row
    
    ' Add the closing HTML table tag.
    result = result & "</table>"
    
    ' Return the assembled HTML table.
    ADVANCED_HTML_TABLE = result
End Function

' Helper Function: RGBToHex
' Description: Converts an RGB color value to a hexadecimal color code.
Private Function RGBToHex(RGBColor As Long) As String
    Dim red As Integer
    Dim green As Integer
    Dim blue As Integer
    
    ' Extract the red, green, and blue components from the RGB color.
    red = RGBColor Mod 256
    green = (RGBColor \ 256) Mod 256
    blue = RGBColor \ 65536
    
    ' Format the hexadecimal color code and return it.
    RGBToHex = Format(Hex(red) & Hex(green) & Hex(blue), "000000")
End Function

