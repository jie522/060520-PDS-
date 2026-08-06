' Maxclaw PDS系統啟動器（GUI 版）
Dim shell : Set shell = CreateObject("WScript.Shell")
Dim fso   : Set fso   = CreateObject("Scripting.FileSystemObject")
Dim base  : base = fso.GetParentFolderName(WScript.ScriptFullName)
Dim pyexe : pyexe = base & "\_python\pythonw.exe"
Dim script: script = base & "\_app\main.py"
Dim url   : url   = "http://127.0.0.1:5088"

' 終止舊的 pythonw 進程
shell.Run "taskkill /F /IM pythonw.exe", 0, True

' 清除 __pycache__（確保載入最新程式碼）
Dim cache : cache = base & "\_app\__pycache__"
If fso.FolderExists(cache) Then fso.DeleteFolder cache, True

' 等待 port 釋放
WScript.Sleep 1200

shell.Run Chr(34) & pyexe & Chr(34) & " " & Chr(34) & script & Chr(34), 0, False
