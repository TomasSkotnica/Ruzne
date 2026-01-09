Option Explicit

Const WshFinished = 1

Dim oExc
'Set oExc = CreateObject("WScript.Shell").Exec("cmd /c dir /s /b")
Set oExc = CreateObject("WScript.Shell").Exec("cmd /c dir /s /b *.vbs | find /c /v """"")
WScript.Echo "A", "start"
Do While True
   If oExc.Status = WshFinished Then
      WScript.Echo "A", "WshFinished"
      Exit Do
   End If
   WScript.Sleep 100
   If Not oExc.Stdout.AtEndOfStream Then WScript.Echo "A in ", oExc.Stdout.ReadLine()
Loop
If Not oExc.Stdout.AtEndOfStream Then WScript.Echo "A out ", oExc.Stdout.ReadAll()
