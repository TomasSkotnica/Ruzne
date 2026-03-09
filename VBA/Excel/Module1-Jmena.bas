Attribute VB_Name = "Module1"
Sub SelecctNames1()
Attribute SelecctNames1.VB_Description = "copy *ra names from source to output sheet"
Attribute SelecctNames1.VB_ProcData.VB_Invoke_Func = "x\n14"
'
' SelecctNames1 Macro
' copy *ra names from source to output sheet
'
' Keyboard Shortcut: Ctrl+x
'
    'Range("A1").Select
    'ActiveCell.FormulaR1C1 = "x"
    'Range("A2").Select
    CopyValuesEndingFromSC "C"
    
End Sub

Sub CopyValuesEndingFromSC(SC As String)

    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim targetRow As Long
    
    Set wsSource = Worksheets("sourceSheet")   'source sheet
    Set wsTarget = Worksheets("outputSheet")   'target sheet
    
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    targetRow = 1
    
    For i = 1 To lastRow
    
        If Right(wsSource.Cells(i, SC).Value, 2) = "ra" Then
        
            wsTarget.Cells(targetRow, 1).Value = wsSource.Cells(i, SC).Value
            targetRow = targetRow + 1
            
        End If
        
    Next i

End Sub
