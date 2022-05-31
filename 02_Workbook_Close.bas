Option Explicit

' Place this code at the ThisWorkbook level within the VBAProject.
' In order for this program to be executed, the procedure "01_Workbook_Open.bas" must first be executed (s. this repository)
' Purpose: Store the logged user activity when closing the workbook.

Private Sub Workbook_BeforeClose(Cancel As Boolean)

Application.DisplayAlerts = False

tbl_logfile.Visible = xlSheetVeryHidden
ThisWorkbook.Save

Application.DisplayAlerts = True

End Sub
