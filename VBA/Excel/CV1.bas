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
Worksheets("Files").Range("D1:D11").Copy _
    Destination:=Worksheets("work").Range("A1")

'Worksheets("work").Range("A1:A11").AdvancedFilter Action:=xlFilterInPlace, Unique:=False
Worksheets("Files").Range("D1:D11").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Worksheets("work").Range("B1"), Unique:=True
    
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
    Dim cell As Range
    Dim lastRow As Long
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    For Each cell In Worksheets(ws).Range(sourceRange)
        If cell.Value <> "" Then
            If Not dict.Exists(cell.Value) Then
                dict.Add cell.Value, cell.Row
            End If
        End If
    Next cell
    
    Set GetDistinct = dict
End Function

Sub PopulatePosts1()
    Worksheets("Files").Range("$A$1:$H$11").AutoFilter ' cancel previous filters
    
    Dim wsPosts As Worksheet
    Worksheets("Posts").Range("A1:G1").Value = Array("Date", "Branch", "Company", "FileName", "Origin", "Folder", "FileId")
    
    Dim customers As Object

    Dim lastRow As Long
    lastRow = Worksheets("Files").Cells(Rows.Count, "A").End(xlUp).Row
    If lastRow = 1 Then ' no records, header line only
        Exit Sub
    End If
    
    Set customers = GetDistinct("Files", "E2:E" & lastRow)
    
    Dim customer As Variant
    For Each customer In customers.Keys
        Debug.Print "--------------"; customer; customers(customer)
        Dim oneRow As Variant
        Selection.AutoFilter
        
        Worksheets("Files").Range("$A$1:$H$11").AutoFilter Field:=5, Criteria1:=customer
        
        For Each oneRow In Worksheets("Files").Range("$A$2:$A$" & lastRow).SpecialCells(xlCellTypeVisible)
            Debug.Print oneRow.Range("A1")
        Next oneRow
        
'        Worksheets("Files").Range("$A$1:$H$" & lastRow).AutoFilter Field:=7, Criteria1:="Microsoft Edge PDF Document"
'        Worksheets("Files").Range("$A$1:$H$" & lastRow).AutoFilter Field:=7, Criteria1:="Microsoft Word Document"
        
        Dim pdfCount
        pdfCount = 0
        Dim docCount
        docCount = 0
        Dim pdfRowNrs As New Collection
        Dim docRowNrs As New Collection
        
        For Each oneRow In Worksheets("Files").Range("$A$2:$A$" & lastRow).SpecialCells(xlCellTypeVisible)
            Debug.Print oneRow.Range("A1")
            If oneRow.Range("G1") = "Microsoft Edge PDF Document" Then
                pdfCount = pdfCount + 1
                pdfRowNrs.Add (oneRow.Row)
            End If
            If oneRow.Range("G1") = "Microsoft Word Document" Then
                docCount = docCount + 1
                docRowNrs.Add (oneRow.Row)
            End If
        Next oneRow
        
        If pdfCount = 1 Then
            Worksheets("Files").Range("$A$1:$H$" & lastRow).AutoFilter Field:=7, Criteria1:="Microsoft Edge PDF Document"
            For Each oneRow In Worksheets("Files").Range("$A$2:$H" & lastRow).SpecialCells(xlCellTypeVisible)
                nextRow = Worksheets("Posts").Cells(Worksheets("Posts").Rows.Count, 1).End(xlUp).Row + 1
                Worksheets("Posts").Cells(nextRow, 1).Value = oneRow.Range("C1")
                Worksheets("Posts").Cells(nextRow, 1).NumberFormat = "yyyy-mm-dd"
                Worksheets("Posts").Cells(nextRow, 2).Value = oneRow.Range("D1")
                Worksheets("Posts").Cells(nextRow, 3).Value = oneRow.Range("E1")
                Worksheets("Posts").Cells(nextRow, 4).Value = oneRow.Range("A1")
                Worksheets("Posts").Cells(nextRow, 5).Value = "not reused"
                Worksheets("Posts").Cells(nextRow, 6).Value = oneRow.Range("F1")
                Worksheets("Posts").Cells(nextRow, 7).Value = oneRow.Range("F1").Row
                Exit For
            Next oneRow
        End If

        Dim rowNumber
        If pdfCount > 1 Then
            Worksheets("Files").Range("$A$1:$H$" & lastRow).AutoFilter Field:=7, Criteria1:="Microsoft Edge PDF Document"
            For Each rowNumber In pdfRowNrs ' oneRow In Worksheets("Files").Range("$A$2:$H" & lastRow).SpecialCells(xlCellTypeVisible)
                oneRow = Worksheets("Files").Range("$A$" & rowNumber & ":$H" & rowNumber).SpecialCells(xlCellTypeVisible)
                Dim fileName, restOfName, companyName As String
                fileName = Worksheets("Files").Cells(rowNumber, 1) 'oneRow.Range("A" & rowNumber)
                restOfName = Replace(fileName, "Tomáš Skotnica CV", "")
                companyName = Trim(Left(restOfName, Len(restOfName) - 4))
                If companyName = "" Then
                    ' find name from motiv or nabidka txt file of the same date, otherwise leave empty
                End If
                If restOfName <> "" Then
                    nextRow = Worksheets("Posts").Cells(Worksheets("Posts").Rows.Count, 1).End(xlUp).Row + 1
                    Worksheets("Posts").Cells(nextRow, 1).Value = Worksheets("Files").Cells(rowNumber, 3)
                    Worksheets("Posts").Cells(nextRow, 1).NumberFormat = "yyyy-mm-dd"
                    Worksheets("Posts").Cells(nextRow, 2).Value = Worksheets("Files").Cells(rowNumber, 4)
                    Worksheets("Posts").Cells(nextRow, 3).Value = companyName
                    Worksheets("Posts").Cells(nextRow, 4).Value = Worksheets("Files").Cells(rowNumber, 1)
                    Worksheets("Posts").Cells(nextRow, 5).Value = "company's original pdf"
                    Worksheets("Posts").Cells(nextRow, 6).Value = Worksheets("Files").Cells(rowNumber, 6)
                    Worksheets("Posts").Cells(nextRow, 7).Value = rowNumber
                End If
            Next rowNumber
        End If
        
        
        If docCount > 0 And pdfCount = 0 Then
            Worksheets("Files").Range("$A$1:$H$" & lastRow).AutoFilter Field:=7, Criteria1:="Microsoft Word Document"
            For Each rowNumber In docRowNrs
                fileName = Worksheets("Files").Cells(rowNumber, 1)
                restOfName = Replace(fileName, "Tomáš Skotnica CV", "")
                companyName = Trim(Left(restOfName, Len(restOfName) - 5))
                If companyName = "" Then
                    companyName = Worksheets("Files").Cells(rowNumber, 5)
                End If
                If restOfName <> "" Then
                    nextRow = Worksheets("Posts").Cells(Worksheets("Posts").Rows.Count, 1).End(xlUp).Row + 1
                    Worksheets("Posts").Cells(nextRow, 1).Value = Worksheets("Files").Cells(rowNumber, 3)
                    Worksheets("Posts").Cells(nextRow, 1).NumberFormat = "yyyy-mm-dd"
                    Worksheets("Posts").Cells(nextRow, 2).Value = Worksheets("Files").Cells(rowNumber, 4)
                    Worksheets("Posts").Cells(nextRow, 3).Value = companyName
                    Worksheets("Posts").Cells(nextRow, 4).Value = Worksheets("Files").Cells(rowNumber, 1)
                    Worksheets("Posts").Cells(nextRow, 5).Value = "company's original doc"
                    Worksheets("Posts").Cells(nextRow, 6).Value = Worksheets("Files").Cells(rowNumber, 6)
                    Worksheets("Posts").Cells(nextRow, 7).Value = rowNumber
                End If
            Next rowNumber
        End If
        
        ' pdfRowNrs must be emptied now by loop, there is no RemoveAll method
        For rowNumber = 1 To pdfRowNrs.Count
            pdfRowNrs.Remove 1    ' Remove the first object each time
        Next rowNumber
        For rowNumber = 1 To docRowNrs.Count
            docRowNrs.Remove 1    ' Remove the first object each time
        Next rowNumber
        
        
        Worksheets("Files").ShowAllData ' cancel previous filters
    Next customer

End Sub

Sub Filt()
Dim ws As Worksheet
Dim lastRow As Long
Dim r As Range

Set ws = Worksheets("Files")
lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

' Apply filter "D2:E" & lastRow
ws.Range("A1").AutoFilter Field:=4, Criteria1:="as"

' Iterate visible (filtered) rows
For Each r In ws.Range("D2:E" & lastRow).SpecialCells(xlCellTypeVisible)
    Debug.Print r.Value, r.Offset(0, 1).Value
Next r

' Remove filter
ws.AutoFilterMode = False

End Sub

Sub ResetFilters()
    Dim lastRow As Long
    lastRow = Worksheets("Files").Cells(Rows.Count, "A").End(xlUp).Row
    Worksheets("Files").Range("$A$1:$H$11").AutoFilter ' no parameters clears all filters, shows all rows
    Worksheets("Files").Range("$A$1:$H$" & lastRow).AutoFilter Field:=7, Criteria1:="Text Document"
End Sub

Sub WriteA3()
    ' Writes value to cell A3 even it is not visible when filter is applied
    Cells(3, 1) = "A3"
End Sub

