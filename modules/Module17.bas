Attribute VB_Name = "Module17"
' Function: colorBox
' Description: Opens a color picker dialog for the user to choose a color from a palette of 160 colors.
' Returns:
'   The RGB color value chosen by the user.
Function colorBox() As Long
    Dim colorDialog As Object
    Set colorDialog = Application.Dialogs(xlDialogEditColor)
    colorDialog.Show
    colorBox = ActiveWorkbook.Colors(56) ' Returns the color chosen by the user
End Function

