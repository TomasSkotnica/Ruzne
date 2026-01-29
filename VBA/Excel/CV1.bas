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
    
    ws.range("A1:H1").Value = Array("File Name", "Full Path", "Last Modified", _
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
Worksheets("Files").range("D1:D11").Copy _
    Destination:=Worksheets("work").range("A1")

'Worksheets("work").Range("A1:A11").AdvancedFilter Action:=xlFilterInPlace, Unique:=False
Worksheets("Files").range("D1:D11").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Worksheets("work").range("B1"), Unique:=True
    
    MsgBox "a"
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


Function GetDistinct(ws As String, sourceRange As String) As Object
    ' this elegant syntax produces duplicits
    'Worksheets("Files").Range("D2:D11").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Worksheets("work").Range("A1"), Unique:=True

    Dim dict As Object
    Dim cell As range
    Dim lastRow As Long
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    For Each cell In Worksheets(ws).range(sourceRange)
        If cell.Value <> "" Then
            If Not dict.Exists(cell.Value) Then
                dict.Add cell.Value, cell.Row
            End If
        End If
    Next cell
    
    Set GetDistinct = dict
End Function

Sub PopulatePosts1()
    Dim wsPosts As Worksheet
    'wsPosts = Worksheets("Posts") 'ThisWorkbook.
    Worksheets("Posts").range("A1:E1").Value = Array("Company", "Date", "Branch", "Folder", "Origin")
    
    Dim level1 As Object

    Dim lastRow As Long
    lastRow = Worksheets("Files").Cells(Rows.Count, "A").End(xlUp).Row
    Set level1 = GetDistinct("Files", "D2:D" & lastRow)
    
    Dim branch As Variant
    For Each branch In level1.Keys
        Debug.Print "--------------"; branch; level1(branch)
        Dim oneRow As Variant
        Selection.AutoFilter
        
        Worksheets("Files").range("$A$1:$H$11").AutoFilter Field:=4, Criteria1:=branch
        
        Dim lastBranchRow
        lastBranchRow = Worksheets("Files").Cells(Rows.Count, 1).End(xlUp).Row
        For Each oneRow In Worksheets("Files").range("$A$2:$A$" & lastBranchRow).SpecialCells(xlCellTypeVisible)
            'Debug.Print oneRow.range("A1")
        Next oneRow
        
        Worksheets("Files").range("$A$1:$H$11").AutoFilter Field:=7, Criteria1:="Microsoft Edge PDF Document"
        
        Dim pdfCount
        pdfCount = 0
        Dim lastPdfRow
        lastPdfRow = Worksheets("Files").Cells(Rows.Count, 1).End(xlUp).Row
        For Each oneRow In Worksheets("Files").range("$A$2:$A$" & lastPdfRow).SpecialCells(xlCellTypeVisible)
            Debug.Print oneRow.range("A1")
            pdfCount = pdfCount + 1
        Next oneRow
        
        If pdfCount = 1 Then
            For Each oneRow In Worksheets("Files").range("$A$13:$H13").SpecialCells(xlCellTypeVisible)
                nextRow = Worksheets("Posts").Cells(Worksheets("Posts").Rows.Count, 1).End(xlUp).Row + 1
                Worksheets("Posts").Cells(nextRow, 1).Value = oneRow.range("A1")
                Worksheets("Posts").Cells(nextRow, 2).Value = oneRow.range("A3")
                Worksheets("Posts").Cells(nextRow, 3).Value = oneRow.range("A4")
                Worksheets("Posts").Cells(nextRow, 4).Value = oneRow.range("A5")
                Worksheets("Posts").Cells(nextRow, 5).Value = "company's original pdf"
            Next oneRow
        End If

        
        Worksheets("Files").ShowAllData
'        Worksheets("Files").range("E2:E11").AdvancedFilter _'            Action:=xlFilterCopy, CopyToRange:=Worksheets("work").range("A1"), CriteriaRange:=range("Criteria")
    Next branch

End Sub

Sub Filt()
Dim ws As Worksheet
Dim lastRow As Long
Dim r As range

Set ws = Worksheets("Files")
lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

' Apply filter "D2:E" & lastRow
ws.range("A1").AutoFilter Field:=4, Criteria1:="as"

' Iterate visible (filtered) rows
For Each r In ws.range("D2:E" & lastRow).SpecialCells(xlCellTypeVisible)
    Debug.Print r.Value, r.Offset(0, 1).Value
Next r

' Remove filter
ws.AutoFilterMode = False

End Sub
