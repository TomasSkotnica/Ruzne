Dim objShell,oExec

Set objShell = wscript.createobject("wscript.shell")
Set oExec = objShell.Exec("calc.exe")
Do While oExec.Status = 0 
    Wscript.Sleep 10000 
Loop
WScript.Echo oExec.Status
WScript.Echo oExec.ProcessID
WScript.Echo oExec.ExitCode 

