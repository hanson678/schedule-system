Set fso = CreateObject("Scripting.FileSystemObject")
Set ws  = CreateObject("WScript.Shell")

' 切换到脚本所在目录
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
ws.CurrentDirectory = scriptDir

' ===== 自动查找 Python =====
pythonExe = ""

' 方法1: 常见安装位置
localAppData = ws.ExpandEnvironmentStrings("%LOCALAPPDATA%")
programFiles = ws.ExpandEnvironmentStrings("%PROGRAMFILES%")
Dim candidates(5)
candidates(0) = localAppData & "\Programs\Python\Python312\python.exe"
candidates(1) = localAppData & "\Programs\Python\Python311\python.exe"
candidates(2) = localAppData & "\Programs\Python\Python310\python.exe"
candidates(3) = programFiles & "\Python312\python.exe"
candidates(4) = programFiles & "\Python311\python.exe"
candidates(5) = programFiles & "\Python310\python.exe"

For Each p In candidates
    If fso.FileExists(p) Then
        pythonExe = p
        Exit For
    End If
Next

' 方法2: 尝试 PATH 中的 python（验证64位）
If pythonExe = "" Then
    On Error Resume Next
    Set exec = ws.Exec("python -c ""import sys; assert sys.maxsize > 2**32; print(sys.executable)""")
    If Err.Number = 0 Then
        pythonExe = Trim(exec.StdOut.ReadAll)
    End If
    On Error GoTo 0
End If

If pythonExe = "" Then
    MsgBox "找不到Python，请先安装Python 3.10+" & vbCrLf & "下载: https://www.python.org/downloads/", vbCritical, "错误"
    WScript.Quit 1
End If

' ===== 读取端口配置 =====
port = "5000"
cfgFile = scriptDir & "\data\config.json"
If fso.FileExists(cfgFile) Then
    Set f = fso.OpenTextFile(cfgFile, 1, False, -1)
    cfgText = f.ReadAll
    f.Close
    Set re = New RegExp
    re.Pattern = """port""\s*:\s*(\d+)"
    re.IgnoreCase = True
    Set matches = re.Execute(cfgText)
    If matches.Count > 0 Then
        port = matches(0).SubMatches(0)
    End If
End If

' ===== 关闭旧进程 =====
ws.Run "cmd /c for /f ""tokens=5"" %a in ('netstat -aon ^| findstr :" & port & " ^| findstr LISTENING') do taskkill /F /PID %a >nul 2>&1", 0, True
WScript.Sleep 500

' ===== 启动应用（静默，无窗口） =====
ws.Run """" & pythonExe & """ app.py", 0, False
WScript.Sleep 1500
ws.Run "http://localhost:" & port
