Option Explicit

' General operation of the Worksheet_Change method.
' Insert this code to the Worksheet where it is to be executed, e.g. 'Sheet1'

Private Sub Worksheet_Change(ByVal Target As Range)

Dim rngCell As Range

For Each rngCell In Target.Cells
  If rngCell.Column = 1 Then
    Select Case rngCell.Value
      Case "Sample A"
        rngCell.Offset(0, 1).Value = "Result A"
      Case "Sample B"
        rngCell.Offset(0, 1).Value = "Result B"
      Case "Sample C"
        rngCell.Offset(0, 1).Value = "Result C"
    End Select
  End If
Next rngCell

' Look up the information stored in the predecessor cell and automatic correction
For Each rngCell In Target.Cells
  If rngCell.Column = 2 Then
    If rngCell.Value <> "Result A" And rngCell.Value <> "Result B" _ 
        And rngCell.Value <> "Result C" Then
      Select Case rngCell.Offset(0, -1).Value
        Case "Sample A"
          rngCell.Value = "Result A"
        Case "Sample B"
          rngCell.Value = "Result B"
        Case "Sample C"
          rngCell.Value = "Result C"
      End Select
    End If
  End If
Next rngCell

End Sub
