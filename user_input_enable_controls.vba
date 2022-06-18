Private Sub btnSubmit_Click()
  Dim cntr As Control
  Dim cmbo As Object
  Dim nmbr As Integer
  
  If optQuestion1A = True Then
    For Each cntr in MultiPage1.Pages(1).Controls
      If TypeName(cntr) = "CheckBox" Then
        If cntr.GroupName = 1 Then cntr.Enabled = True
      End If
    Next
    
    For nmbr = 1 To 5
      Set cmbo = MultiPage1.Pages("pgCheckpoints").Controls("ComboBox" & nmbr)
      With cmbo
        .Enabled = True
      End With
    Next
  
  Else
    For Each cntr in MultiPage1.Pages(1).Controls
      If TypeName(cntr) = "CheckBox" Then
        If cntr.GroupName = 1 Then cntr.Enabled = False
      End If
    Next
    
    For nmbr = 1 To 5
      Set cmbo = MultiPage1.Pages("pgCheckpoints").Controls("ComboBox" & nmbr)
      With Combo
        .Enabled = False
      End With
    Next

  End If
 
  MultiPage1.Pages("pgCheckpoints").Enabled = True
  MultiPage1.Value = 1
  
End Sub
