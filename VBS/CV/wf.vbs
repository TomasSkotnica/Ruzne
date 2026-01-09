
Dim shell, exec, result

Set shell = CreateObject("WScript.Shell")
WScript.Echo "start"

Set exec = shell.Exec("cmd /c dir /s /b *.vbs | find /c /v """" > countOf.txt")

' this doesn't work:
'Set exec = shell.Exec("cmd /c dir /s /b *.vbs | find /c /v """"")
'Do While exec.Status = 0 :Wscript.Sleep 2000 :Loop
'result = Trim(exec.StdOut.ReadAll)

WScript.Quit
'''''''''''''''''''''

On Error Resume Next

Dim fso, file

Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile("c:\work\Ruzne\VBS\CV\test.txt", 8, True) : file.WriteLine :
file.WriteLine "Script started " & Now

If Err.Number <> 0 Then
    WScript.Echo "Error writing file: " & Err.Description
    WScript.Quit 1
End If

file.WriteLine "Result of cmd /c dir /s /b *.vbs ^| find /c /v """" is " & result
file.Close

WScript.Quit


