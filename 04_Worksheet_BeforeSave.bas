Option Explicit

Private Sub Workbook_BeforeSave _
(ByVal SaveAsUI As Boolean, Cancel As Boolean)

Dim lngNextCell As Long
Dim wksSheet As Worksheet

Set wksSheet = tbl_Logfile

lngNextCell = wksSheet.Cells(wksSheet.Rows.Count, 1).End(xlUp).Row + 1

With wksSheet
  .Cells(1, 1).Value = "User"
  .Cells(1, 2).Value = "Operation"
  .Cells(1, 3).Value = "Date, Time"
  .Cells(lngNextCell, 1).Value = Environ("username")
  .Cells(lngNextCell, 2).Value = "saved changes"
  .Cells(lngNextCell, 3).Value = Now()
  .Columns.AutoFit
End With

End Sub
