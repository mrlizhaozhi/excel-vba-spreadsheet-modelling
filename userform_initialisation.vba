Private Sub UserForm_Initialize()
  MultiPage1.Pages("pgCheckpoints").Enabled = False
  btnSubmit.Enabled = False
  btnSubmit.BackColor = &H8000000F
  
  Dim cntr As Control
  For Each cntr In MultiPage1.Pages(1).Controls
    If TypeName(cntr) = "CheckBox" Then cntr.Enabled = False
    If TypeName(cntr) = "ComboBox" Then
      cntr.List = Array("Completed", "Checked with Errors", "Missing")
      cntr.Enabled = False
    End If
  Next
End Sub
