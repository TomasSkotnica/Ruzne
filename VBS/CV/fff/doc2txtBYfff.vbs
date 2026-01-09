' cscript doc2txtBYfff.vbs c:\Users\t\Documents\CV\ c:\temp\doc2txt\dest\

' the first arg is needed to recognize relative part of path
' the second arg specifies folder with input list of files to process as well as target dir

' fff in script name means files from file
' by command: dir /s /b c:\Users\t\Documents\CV\*.docx > c:\temp\doc2txt\dest\list.txt
' to get number of files: dir /s /b c:\Users\t\Documents\CV\*.docx | find /c ".docx"
' or better find /c /v "" (when dir /b is used)

If WScript.Arguments.Count < 2 Then
  WScript.Echo "Usage: " + WScript.ScriptName + " <source folder> <destination folder>"
  WScript.Quit 1
End If

Set args = Wscript.Arguments
Wscript.Echo WScript.ScriptFullName
Wscript.Echo "List of arguments:"
For Each arg In args
  Wscript.Echo arg
Next
Wscript.Echo "Script starts:"

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

sourceRoot = WScript.Arguments(0)
If Not fso.FolderExists(sourceRoot) Then
    WScript.Echo "Folder not found: " & sourceRoot
    WScript.Quit 1
End If

destRoot = WScript.Arguments(1)
If Not fso.FolderExists(destRoot) Then
    WScript.Echo "Folder not found: " & destRoot
    WScript.Quit 1
End If

if Right(destRoot, 1) <> "\" Then destRoot = destRoot & "\"
WScript.Echo "po uprave: " & destRoot

Set logf = fso.OpenTextFile(destRoot & "test.txt", 8, True)
logf.WriteLine ""
logf.WriteLine "Script started at " & Now

Set objWord = CreateObject("Word.Application")

Set list = fso.OpenTextFile(destRoot & "list.txt", 1)

Do Until list.AtEndOfStream
    oneDocFile = list.ReadLine
    Wscript.Echo oneDocFile
    name = Right(oneDocFile, Len(oneDocFile) - InStrRev(oneDocFile, "\"))
    WScript.Echo name

    sourceFolder = Left(oneDocFile, InStrRev(oneDocFile, "\"))
    Wscript.Echo sourceFolder
    relat = Right(sourceFolder, Len(sourceFolder) - Len(sourceRoot))
    WScript.Echo relat

    destFolder = destRoot & relat
    strNewName =  destFolder & Left(name, Len(name) - 4) & "txt"
    WScript.Echo "new file = " &strNewName

    If Not fso.FolderExists(destFolder) Then
        Call CreateFolder(destRoot, relat)
    End If

    ' convert to txt
    ' On Error Resume Next, supresses error message writing to command line
    ' then error can be handled by If Err.Number <> 0
    ' when it is not used, it writes the message and ends the script
    On Error Resume Next
    Set objDoc = objWord.Documents.Open(oneDocFile)
    objDoc.SaveAs strNewName, 2
    objDoc.Close

    If fso.FileExists(strNewName) Then
        logf.WriteLine strNewName & " OK"
    Else
        logf.WriteLine strNewName & " failed"
    End If

Loop

list.Close

WScript.Echo "Quiting Word"
objWord.Quit

logf.WriteLine "Script ended at " & Now
logf.WriteLine ""
logf.Close


'''''''''''''''''''''''''''''''''
Sub CreateFolder(where, what)
    WScript.Echo where
    Dim arr
    arr = Split(what, "\")
    subCounter = ""
    Dim i
    For i = 0 To UBound(arr)
        WScript.Echo arr(i)
        if arr(i) <> "" Then 
            subCounter = subCounter & arr(i) & "\"
            WScript.Echo subCounter
            fso.CreateFolder(where & subCounter)
        End If
    Next
End Sub