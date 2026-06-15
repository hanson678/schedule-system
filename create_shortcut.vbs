Set ws  = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' 自动获取脚本所在目录（即排期系统文件夹）
projectDir = fso.GetParentFolderName(WScript.ScriptFullName)
desktop = ws.SpecialFolders("Desktop")

' 创建桌面快捷方式 → 指向一键启动.vbs
Set sc = ws.CreateShortcut(desktop & "\ZURU" & ChrW(25490) & ChrW(26399) & ChrW(31995) & ChrW(32479) & ".lnk")
sc.TargetPath = projectDir & "\" & ChrW(19968) & ChrW(38190) & ChrW(21551) & ChrW(21160) & ".vbs"
sc.WorkingDirectory = projectDir
sc.IconLocation = "C:\Windows\System32\shell32.dll,165"
sc.Save
