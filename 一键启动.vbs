Set ws = CreateObject("WScript.Shell")
ws.CurrentDirectory = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
' 先关闭所有旧的app.py进程（端口5000），避免多个进程冲突
ws.Run "cmd /c taskkill /F /FI ""WINDOWTITLE eq ZURU*"" >nul 2>&1 & for /f ""tokens=5"" %a in ('netstat -aon ^| findstr :5000 ^| findstr LISTENING') do taskkill /F /PID %a >nul 2>&1", 0, True
WScript.Sleep 500
ws.Run """C:\Users\Administrator\AppData\Local\Programs\Python\Python312\python.exe"" app.py", 0, False
WScript.Sleep 1500
ws.Run "http://localhost:5000"
