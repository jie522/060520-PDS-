' 製令查詢系統啟動器（GUI 版）
Dim shell : Set shell = CreateObject("WScript.Shell")
Dim fso   : Set fso   = CreateObject("Scripting.FileSystemObject")
Dim base  : base = fso.GetParentFolderName(WScript.ScriptFullName)
Dim pyexe : pyexe = base & "\_python\pythonw.exe"
Dim script: script = base & "\_app\main.py"
Dim url   : url   = "http://127.0.0.1:5088"

' 終止本安裝的舊 pythonw 進程（靜默，忽略錯誤）
shell.Run "taskkill /F /IM pythonw.exe", 0, True

' 等待 port 釋放
WScript.Sleep 1200

' 切換工作目錄到 _app，再啟動 Flask
Dim appDir : appDir = base & "\_app"
shell.CurrentDirectory = appDir
shell.Run Chr(34) & pyexe & Chr(34) & " main.py", 0, False

' 等待 Flask 就緒
WScript.Sleep 3000

' 以 Edge App 模式開啟視窗
Dim edge : edge = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
If Not fso.FileExists(edge) Then
    edge = "C:\Program Files\Microsoft\Edge\Application\msedge.exe"
End If
shell.Run Chr(34) & edge & Chr(34) & " --app=" & url & " --window-size=1400,900", 1, False
