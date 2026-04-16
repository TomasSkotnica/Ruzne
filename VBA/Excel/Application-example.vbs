Dim xl
Set xl = CreateObject("Excel.Application")

xl.Visible = True
xl.Workbooks.Add

Set myObject = xl.Workbooks(1)
If myObject.Application.Value = "Microsoft Excel" Then 
 MsgBox "This is an Excel Application object." 
Else 
 MsgBox "This is not an Excel Application object." 
End If