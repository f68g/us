Dim shell, tempFolder
Set shell = CreateObject("WScript.Shell")

' Lấy đúng thư mục Temp thật của hệ thống
tempFolder = shell.ExpandEnvironmentStrings("%TEMP%") & "\"

' Đường dẫn tới 1.vbs và 2.vbs trong Temp
file1 = """" & tempFolder & "1.vbs" & """"
file2 = """" & tempFolder & "2.vbs" & """"

' Chạy 1.vbs trước, chờ nó chạy xong
shell.Run "wscript.exe " & file1, 0, True

' Sau đó chạy 2.vbs
shell.Run "wscript.exe " & file2, 0, False
