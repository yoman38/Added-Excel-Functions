Attribute VB_Name = "Module16"
' Function: colorToHexa
' Description: Returns the hexadecimal value of a color from a Color value.
' Parameters:
'   colorValue: The Color value to convert to hexadecimal.
' Returns:
'   The hexadecimal representation of the color value.
Function colorToHexa(ByVal colorValue As Long) As String
    Dim redVal As Integer
    Dim greenVal As Integer
    Dim blueVal As Integer
    
    ' Extract the RGB components from the color value.
    redVal = colorValue Mod 256
    greenVal = (colorValue \ 256) Mod 256
    blueVal = (colorValue \ 65536) Mod 256
    
    ' Convert the RGB components to hexadecimal and concatenate them.
    colorToHexa = "#" & Right("0" & Hex(redVal), 2) & Right("0" & Hex(greenVal), 2) & Right("0" & Hex(blueVal), 2)
End Function

' Function: hexaToColor
' Description: Returns the Color value from a hexadecimal color representation.
' Parameters:
'   hexaColor: The hexadecimal color value (e.g., "#00ff00") to convert to Color.
' Returns:
'   The Color value of the hexadecimal color, or -1 in case of an error.
Function hexaToColor(ByVal hexaColor As String) As Long
    On Error Resume Next
    ' Remove the "#" character if it exists.
    If Left(hexaColor, 1) = "#" Then
        hexaColor = Mid(hexaColor, 2)
    End If
    
    ' Convert the hexadecimal components to decimal (RGB) values.
    Dim redVal As Integer
    Dim greenVal As Integer
    Dim blueVal As Integer
    
    redVal = CLng("&H" & Mid(hexaColor, 1, 2))
    greenVal = CLng("&H" & Mid(hexaColor, 3, 2))
    blueVal = CLng("&H" & Mid(hexaColor, 5, 2))
    
    ' Check if conversion was successful.
    If Err.Number <> 0 Then
        hexaToColor = -1
    Else
        ' Combine the RGB values to get the final Color value.
        hexaToColor = RGB(redVal, greenVal, blueVal)
    End If
    On Error GoTo 0
End Function

