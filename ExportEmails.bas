Attribute VB_Name = "Module1"
Function EscapeQuotesForJSON(inputString As String) As String
    ' Replace backslashes (\) with double backslashes (\\)
    inputString = Replace(inputString, "\", "\\")
    ' Replace quotes (") with escaped quotes (\")
    EscapeQuotesForJSON = Replace(inputString, """", "\""")
End Function

Sub GetEmails()
   ' On Error GoTo ErrorHandler
    Dim OutlookApp As Outlook.Application
    Dim OutlookNamespace As NameSpace
    Dim Inbox As MAPIFolder
    Dim Email As Object
    Dim i As Integer
    Dim FileSystem As Object
    Dim JsonFile As Object
    Dim JsonString

    Set OutlookApp = New Outlook.Application
    Set OutlookNamespace = OutlookApp.GetNamespace("MAPI")
    Set Inbox = OutlookNamespace.GetDefaultFolder(olFolderInbox)
    Set FileSystem = CreateObject("Scripting.FileSystemObject")
    Set JsonFile = FileSystem.CreateTextFile("Here you must set the proper path to outlook_topic_summarizer\emails.json", True, True)

    JsonString = "["
    i = 0
    For Each Email In Inbox.Items
        If i >= 22 Then Exit For
        JsonString = JsonString & "{""title"": """ & EscapeQuotesForJSON(Email.Subject) & """, ""body"": """ & EscapeQuotesForJSON(Email.Body) & """},"
        i = i + 1
    Next Email
    ' Remove trailing comma and close the JSON array
    JsonString = Left(JsonString, Len(JsonString) - 1) & "]"
    JsonFile.Write JsonString
    JsonFile.Close
    MsgBox "All done!"
    Set Email = Nothing
    Set Inbox = Nothing
    Set OutlookNamespace = Nothing
    Set OutlookApp = Nothing
    Set FileSystem = Nothing
    Set JsonFile = Nothing
End Sub
