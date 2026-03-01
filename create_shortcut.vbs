Set ws = CreateObject("WScript.Shell")
desktop = ws.SpecialFolders("Desktop")
Set sc = ws.CreateShortcut(desktop & "\" & "ZURU" & ChrW(25490) & ChrW(26399) & ChrW(31995) & ChrW(32479) & ".lnk")
sc.TargetPath = desktop & "\" & ChrW(25490) & ChrW(26399) & ChrW(31995) & ChrW(32479) & "\" & ChrW(19968) & ChrW(38190) & ChrW(21551) & ChrW(21160) & ".vbs"
sc.WorkingDirectory = desktop & "\" & ChrW(25490) & ChrW(26399) & ChrW(31995) & ChrW(32479)
sc.IconLocation = "C:\Windows\System32\shell32.dll,165"
sc.Save
