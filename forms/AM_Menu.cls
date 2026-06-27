Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub CloseMe_Click()
  '  DoCmd.OpenForm "FAMainMenu"
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub GoHome_Click()
 DoCmd.OpenForm "FAMainMenu"
End Sub

Private Sub SerialNumberEntry_Click()
DoCmd.OpenForm "UnitReverseLookUpForm"
End Sub
