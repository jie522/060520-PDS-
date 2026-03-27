' 製令查詢系統啟動器（GUI 版）
Dim shell : Set shell = CreateObject("WScript.Shell")
Dim fso   : Set fso   = CreateObject("Scripting.FileSystemObject")
Dim base  : base = fso.GetParentFolderName(WScript.ScriptFullName)
Dim pyexe : pyexe = base & "\_python\pythonw.exe"
Dim script: script = base & "\_app\main.py"
shell.Run Chr(34) & pyexe & Chr(34) & " " & Chr(34) & script & Chr(34), 0, False
