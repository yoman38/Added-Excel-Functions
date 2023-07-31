Attribute VB_Name = "Module15"
' Function: isEmail
' Description: Returns True if the string is a valid email address or False otherwise.
' Parameters:
'   emailString: The string to check if it's a valid email address.
' Returns:
'   True if the string is a valid email address, False otherwise.
Function isEmail(ByVal emailString As String) As Boolean
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    ' Regular expression pattern for email validation.
    regEx.Pattern = "^[\w\.-]+@[a-zA-Z\d\.-]+\.[a-zA-Z]{2,}$"
    
    isEmail = regEx.Test(emailString)
End Function

' Function: isUrl
' Description: Returns True if the string is a valid URL or False otherwise.
' Parameters:
'   urlString: The string to check if it's a valid URL.
' Returns:
'   True if the string is a valid URL, False otherwise.
Function isUrl(ByVal urlString As String) As Boolean
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    ' Regular expression pattern for URL validation.
    regEx.Pattern = "^(https?|ftp)://[^\s/$.?#].[^\s]*$"
    
    isUrl = regEx.Test(urlString)
End Function

' Sub: mail
' Description: Sends an email using an email solution compatible with all email addresses.
' Parameters:
'   emailTo: The email address of the recipient.
'   subject: The subject of the email.
'   body: The body of the email.
'   Optional attachmentFilePath As String: The file path of the attachment (if any).
Sub mail(ByVal emailTo As String, ByVal subject As String, ByVal body As String, Optional ByVal attachmentFilePath As String)
    ' Code to send the email goes here.
    ' Implementation depends on the email solution being used.
    ' Replace the code below with the actual code to send the email.
    ' For example, you can use a web API or a third-party email service.
End Sub

' Function: htmlCodePage
' Description: Retrieves the HTML code of a web page.
' Parameters:
'   url: The URL of the web page to retrieve the HTML code from.
' Returns:
'   The HTML code of the web page, or -1 in case of an error.
Function htmlCodePage(ByVal url As String) As Variant
    On Error Resume Next
    Dim xmlHttp As Object
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
    xmlHttp.Open "GET", url, False
    xmlHttp.send
    
    If xmlHttp.Status = 200 Then
        htmlCodePage = xmlHttp.responseText
    Else
        htmlCodePage = -1
    End If
    
    Set xmlHttp = Nothing
    On Error GoTo 0
End Function

' Function: internet
' Description: Returns True if the computer is connected to the internet or False otherwise.
' Parameters: None
' Returns:
'   True if connected to the internet, False otherwise.
Function internet() As Boolean
    On Error Resume Next
    internet = (InStr(1, CreateObject("msxml2.xmlhttp").Open("GET", "https://www.google.com", False), "200 OK") > 0)
    On Error GoTo 0
End Function

' Function: linkOpen
' Description: Opens a web page link and returns True. Returns False if the link cannot be opened.
' Parameters:
'   linkUrl: The URL of the web page link to open.
' Returns:
'   True if the link is successfully opened, False otherwise.
Function linkOpen(ByVal linkUrl As String) As Boolean
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink Address:=linkUrl
    linkOpen = (Err.Number = 0)
    On Error GoTo 0
End Function

