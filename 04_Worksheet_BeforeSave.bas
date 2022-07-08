Option Explicit

' Place this code at the ThisWorkbook level within the VBAProject.
' Purpose: Log user activity when saving the workbook.
' Note: The table tbl_logfile_change must be created first.

Private Sub Workbook_BeforeSave _
(ByVal SaveAsUI As Boolean, Cancel As Boolean)

Dim lngNextCell As Long
Dim wksSheet As Worksheet

Set wksSheet = tbl_logfile

With wksSheet
    .Visible = xlSheetVisible
    .Cells(1, 1).Value = "Date"
    .Cells(1, 2).Value = "Time"
    .Cells(1, 3).Value = "Username"
    .Cells(1, 4).Value = "Hostname"
    .Cells(1, 5).Value = "Operation"
    lngNextCell = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
    .Cells(lngNextCell, 1).Value = Date
    .Cells(lngNextCell, 2).Value = Time
    .Cells(lngNextCell, 3).Value = Environ("username")
    .Cells(lngNextCell, 4).Value = Environ("computername")
    .Cells(lngNextCell, 5).Value = "saved changes"
    .Columns.AutoFit
    .Visible = xlSheetVeryHidden
End With

End Sub
