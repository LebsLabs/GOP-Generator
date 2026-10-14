Attribute VB_Name = "Module2"
Sub TestLocalFolder()

    Dim folderPath As String
    Dim enteredPassword As String
    Dim correctPassword As String
    
    'set password
    correctPassword = "DEMO_PASSWORD"
    
    'ask user for password
    enteredPassword = InputBox("Please enter the password to continue:", _
    "Password Required")
    
    'check for password
    If enteredPassword <> correctPassword Then
    MsgBox "Incorrect password. Access denied.", vbCritical, "Access Denied"
    Exit Sub
    End If
    
    'continue with macro
    folderPath = Environ$("USERPROFILE") & "\Documents\GOP-Generator"

    If Dir(folderPath, vbDirectory) <> "" Then
        MsgBox "The local folder was found."
    Else
        MsgBox "The local folder was not found."
    End If

End Sub
