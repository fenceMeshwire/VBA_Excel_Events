Option Explicit

' Place this code at the ThisWorkbook level within the VBAProject.
' Purpose: Log user activity when opening the workbook.
' Note: The table tbl_logfile must be created first.

Private Sub Workbook_Open()

Dim lngNextCell As Long
Dim strUsername As String

strUsername = "Your_Username"

Application.DisplayAlerts = False

' Monitor user activity
With tbl_logfile
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
    .Cells(lngNextCell, 5).Value = "opened workbook"
    .Visible = xlSheetVeryHidden
    ThisWorkbook.Save
End With

Select Case Environ("username")
    Case strUsername
        MsgBox "You are granted access to the logfile."
        tbl_logfile.Visible = xlSheetVisible
        Exit Sub
    Case Else
        tbl_logfile.Visible = xlSheetVeryHidden
End Select

Application.DisplayAlerts = True

End Sub
