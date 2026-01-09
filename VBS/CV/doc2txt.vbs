' cscript doc2txt.vbs c:\Users\t\Documents\CV\doc2txt\AS\ c:\temp\doc2txt\dest\
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

Dim fso, sourceFolderObject, destRootObject
Set fso = CreateObject("Scripting.FileSystemObject")

sourceFolder = WScript.Arguments(0)
If Not fso.FolderExists(sourceFolder) Then
    WScript.Echo "Folder not found: " & sourceFolder
    WScript.Quit 1
End If

destRoot = WScript.Arguments(1)
If Not fso.FolderExists(destRoot) Then
    WScript.Echo "Folder not found: " & destRoot
    WScript.Quit 1
End If

set sourceFolderObject = fso.GetFolder(sourceFolder)
sourceFolderLength = Len(sourceFolderObject.Path)
if Right(destRoot, 1) <> "\" Then destRoot = destRoot & "\"
WScript.Echo "po uprave: " & destRoot

Set logf = fso.OpenTextFile(destRoot & "test.txt", 8, True)

Set objWord = CreateObject("Word.Application")
Call ListDocxFiles(sourceFolder)

WScript.Echo "ListDocxFiles ended"
WScript.Echo "Quiting Word"
objWord.Quit
logf.Close
' end of main

''''''''''''''''''''''''''''''''''''
Sub ListDocxFiles(folderPath)
    Dim folder, subfolder, file
    Set folder = fso.GetFolder(folderPath)

    For Each file in folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "docx" Then
            Wscript.Echo
            WScript.Echo file.Path
            relat  = Right(folder.Path, Len(folder.Path) - sourceFolderLength)
            destFolder = destRoot & relat & "\"
            'WScript.Echo destFolder
            If Not fso.FolderExists(destFolder) Then
                'WScript.Echo "Creating: " & destFolder
                fso.CreateFolder(destFolder)
            End If

            strNewName = destFolder & Left(file.Name, Len(file.Name) - 4) & "txt"
            WScript.Echo strNewName

            ' On Error Resume Next, supresses error message writing to command line
            ' then error can be handled by If Err.Number <> 0
            ' when it is not used, it writes the message and ends the script
            On Error Resume Next
            Set objDoc = objWord.Documents.Open(file.Path)
            If Err.Number <> 0 Then
                WScript.Echo "My error message is: " & Err.Number & " " & Err.Source & " " & Err.Description
                'MsgBox("Err.Number is " & Err.Number)
                Err.Clear() ' Clear Err object fields.
            End If            
            objDoc.SaveAs strNewName, 2
'            WScript.Echo "Press any key to continue ..." : WScript.StdIn.ReadLine
            objDoc.Close

            If fso.FileExists(strNewName) Then
                logf.WriteLine strNewName & " OK"
            Else
                logf.WriteLine strNewName & " failed"
            End If

        End If
    Next

    For Each subfolder in folder.SubFolders
        Call ListDocxFiles(subfolder.Path)
    Next
End Sub

