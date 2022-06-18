Private Sub btnCount_Click()
  Dim optBtn As Control
  Dim count As Integer
  
  For Each optBtn in MultiPage1.Pages("pgQuestions").Controls
    If TypeName(optBtn) = "OptionButton" Then
      If optBtn = True Then
        count = count + 1
      End If
    End If
  Next
  
  If count = 7 Then
    btnSubmit.Enabled = True
    btnSubmit.BackColor = &HFF&
  End If
End Sub
