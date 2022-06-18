Private Sub btnSubmit2_Click()
  Worksheets("GuaranteePersonal").Activate
  Dim emptyRow As Long
  
  emptyRow = WorksheetFunction.CountA(Range("B:B")) + 1
  
  If optQuestion1A = True Then
    Cells(emptyRow, 2).Value = "A"
  ElseIf optQuestion1B = True Then
    Cells(emptyRow, 2).Value = "B"
  End If
End Sub
