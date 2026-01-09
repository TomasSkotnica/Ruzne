'On Error Resume Next
sourceRoot = "c:\Users\t\Documents\CV\"
destRoot = "c:\temp\doc2txt\dest\"

Wscript.Echo sourceRoot
Wscript.Echo destRoot
Wscript.Echo "--------------"
Dim fso, sourceFolderObject, destRootObject
Set fso = CreateObject("Scripting.FileSystemObject")
set sourceFolderObject = fso.GetFolder(sourceRoot)
sourceFolderLength = Len(sourceFolderObject.Path)


oneDocFile = "c:\Users\t\Documents\CV.doc"
'D(oneDocFile) ' this makes failure - file is not in source root
oneDocFile = "c:\Users\t\Documents\CV\CV.doc"
D(oneDocFile)
oneDocFile = "c:\Users\t\Documents\CV\CV.docx"
D(oneDocFile)
oneDocFile = "c:\Users\t\Documents\CV\A\CV.docx"
D(oneDocFile)
oneDocFile = "c:\Users\t\Documents\CV\A\B\CV.docx"
D(oneDocFile)


Wscript.Quit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub D(oneDocFile)
Wscript.Echo 
Wscript.Echo 
Wscript.Echo oneDocFile
sourceFolder = Left(oneDocFile, InStrRev(oneDocFile, "\"))
Wscript.Echo sourceFolder
ll = InStrRev(oneDocFile, "\")
'WScript.Echo sourceFolderLength
'WScript.Echo Len(oneDocFile)
'Wscript.Echo ll
name = Right(oneDocFile, Len(oneDocFile) - InStrRev(oneDocFile, "\"))
WScript.Echo name

relat = Right(sourceFolder, Len(sourceFolder) - Len(sourceRoot))
WScript.Echo relat

'WScript.Echo
strNewName = destRoot & relat & Left(name, Len(name) - 4) & "txt"
WScript.Echo "res = " &strNewName
End Sub
