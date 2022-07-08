Option Explicit

' Place this code at the ThisWorkbook level within the VBAProject.
' In order to be executed, the WorkSheet tbl_logfile has to exist and 
' the procedure "01_Workbook_Open.bas" must be implemented (s. this repository)

' Purpose: Finally hide the WorkSheet tbl_logfile if made visible by an administrator.

Private Sub Workbook_BeforeClose(Cancel As Boolean)

Application.DisplayAlerts = False

If tbl_logfile.Visible = xlSheetVisible Then
  tbl_logfile.Visible = xlSheetVeryHidden
End If

ThisWorkbook.Save

Application.DisplayAlerts = True

End Sub
