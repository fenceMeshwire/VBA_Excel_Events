Option Explicit

' Place this code at the ThisWorkbook level within the VBAProject.
' Purpose: Log user activity when opening the workbook.

Private Sub Workbook_Open()

Dim lngCellFree As Long
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
    lngCellFree = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
    .Cells(lngCellFree, 1).Value = Date
    .Cells(lngCellFree, 2).Value = Time
    .Cells(lngCellFree, 3).Value = Environ("username")
    .Cells(lngCellFree, 4).Value = Environ("computername")
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
