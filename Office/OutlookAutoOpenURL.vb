Option Explicit

' Make user of Windows API to open URLs in the default browser
Private Declare PtrSafe Function ShellExecute _
  Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hWnd As Long, _
    ByVal Operation As String, _
    ByVal Filename As String, _
    Optional ByVal Parameters As String, _
    Optional ByVal Directory As String, _
    Optional ByVal WindowStyle As Long = vbMinimizedFocus _
  ) As Long
  
Private WithEvents InboxItems As Outlook.Items

Private Sub Application_Startup()
    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")
    Set InboxItems = ns.GetDefaultFolder(olFolderInbox).Items
End Sub

Private Sub InboxItems_ItemAdd(ByVal Item As Object)
    On Error Resume Next ' In case of errors, skip to next line (use with caution)
    
    If TypeName(Item) = "MailItem" Then
        Call ProcessMailItem(Item)
    End If
End Sub

'Occurs when incoming emails arrive in Inbox

Sub ProcessMailItem(ByVal mail As Outlook.MailItem)
    Dim bodyText As String
    Dim regex As Object
    Dim lSuccess As Long
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Define the URL pattern here
    regex.Pattern = "https?:\/\/[^\s]*secure\.force\.com[^\s]*"
    regex.Global = False ' Set to True if you want to find all matches

    bodyText = mail.Body
    
    ' Search for the URL
    If regex.Test(bodyText) Then
        Dim matches As Object
        Set matches = regex.Execute(bodyText)
        Dim url As String
        url = matches(0).Value
        Debug.Print url
        ' Open the URL in the default browser
        lSuccess = ShellExecute(0, "Open", url)
    End If
End Sub
