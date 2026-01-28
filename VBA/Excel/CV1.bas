Attribute VB_Name = "CV1"
Option Explicit

Dim fso As Object
Dim nextRow As Long
Dim rootPath As String
Dim dateFrom As Date




Sub ListFilesFromFolder(rp As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
       
    rootPath = LCase(rp)
    Do While Right(rootPath, 1) = "\"
        rootPath = Left(rootPath, Len(rootPath) - 1)
    Loop
    
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ws.Range("A1:H1").Value = Array("File Name", "Full Path", "Last Modified", _
                                    "Folder Level 1", "Folder Level 2", "Folder Level 3", "File Type", "Content Type")
    Set fso = CreateObject("Scripting.FileSystemObject")

    ProcessFolder rootPath, rootPath, ws
End Sub

Sub ProcessFolder(ByVal currentPath As String, ByVal rootPath As String, ws As Worksheet)
    Dim folder As Object
    Dim subFolder As Object
    Dim file As Object
    
    Set folder = fso.GetFolder(currentPath)
    
    For Each file In folder.Files
        If DateWithinSpecScope(dateFrom, file) Then WriteFileRow file, rootPath, ws
    Next file
    
    For Each subFolder In folder.SubFolders
        ProcessFolder subFolder.Path, rootPath, ws
    Next subFolder
End Sub
Function DateWithinSpecScope(dateFrom As Date, file As Object) As Boolean
    Dim dd As String
    dd = file.DateLastModified
    
    DateWithinSpecScope = file.DateLastModified >= dateFrom
End Function


Sub WriteFileRow(file As Object, rootPath As String, ws As Worksheet)
    Dim relativePath As String
    Dim folders() As String
    Dim contentType As String
    Dim deeperLevels As String
    
    ' Get path relative to root
    relativePath = Replace(LCase(file.ParentFolder.Path), rootPath, "")
    
    
    ws.Cells(nextRow, 1).Value = file.Name
    ws.Cells(nextRow, 2).Value = file.Path
    ' Display a date by using the short date format specified in your computer's regional settings.
    ws.Cells(nextRow, 3).Value = file.DateLastModified 'FormatDateTime(file.DateLastModified, vbShortDate)
    ws.Cells(nextRow, 3).NumberFormat = "yyyy-mm-dd"
    
    If relativePath <> "" Then
        If Left(relativePath, 1) = "\" Then relativePath = Mid(relativePath, 2)
        folders = Split(relativePath, "\")
    
        ws.Cells(nextRow, 4).Value = IIf(UBound(folders) >= 0, folders(0), "")
        If UBound(folders) >= 1 Then ws.Cells(nextRow, 5).Value = folders(1)
        
        If UBound(folders) >= 2 Then ws.Cells(nextRow, 6).Value = folders(2)
        If UBound(folders) >= 3 Then
            deeperLevels = Mid(relativePath, Len(folders(0)) + Len(folders(1)) + 3)
            ws.Cells(nextRow, 6).Value = deeperLevels
        End If
    End If
   
    ws.Cells(nextRow, 7).Value = file.Type
    
    contentType = "jiny"
    If InStr(1, UCase(file.Name), "CV", vbTextCompare) > 0 Then contentType = "CV"
    If InStr(1, LCase(file.Name), "motiv", vbTextCompare) > 0 Then contentType = "motiv"
    If InStr(1, LCase(file.Name), "nabidka", vbTextCompare) > 0 Then contentType = "nabidka"
    ws.Cells(nextRow, 8).Value = contentType
    
    nextRow = nextRow + 1
End Sub

Sub Test()
    Dim from As Date
    
    from = DateValue("2026-01-20")
    MsgBox from
End Sub
Sub ShortcutRefreshCVlist()
    CVInputForm.Show
End Sub

Sub AktualizaceDebug()
    dateFrom = DateValue("2022-01-20")
    ListFilesFromFolder LCase("c:\Users\t\Documents\CV\_ostatni\")
    MsgBox "Done", vbInformation
End Sub

Sub AktualizaceFromGUI()
    Dim rootPathGUI As String
    rootPathGUI = LCase(CVInputForm.boxRootFolder.Value)
    dateFrom = DateValue(CVInputForm.boxDateFrom.Value)
    ListFilesFromFolder rootPathGUI
    MsgBox "Done", vbInformation
End Sub

