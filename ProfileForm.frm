
'==============================================================================================
' Profile Viewer
'
' Author: Ben Pogrund
'
'==============================================================================================

Public ViewIndex As Integer

Private Sub chkProfSave_Click()
  chkProfSave.Value = 0
  Send0x27 ViewIndex
End Sub

Private Sub chkProfClose_Click()
  chkProfClose.Value = 0
  ProfileForm.Hide
End Sub

