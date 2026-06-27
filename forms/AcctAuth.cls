Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub AcctPW_AfterUpdate()
  Call AcctSubmitBtn_Click
End Sub

Private Sub AcctPW_Enter()
  Call AcctSubmitBtn_Click
End Sub

Private Sub AcctSubmitBtn_Click()
   If IsNull(Me.AcctPW) Or Me.AcctPW <> "2014AcctPW" Then
       MsgBox "Invalid Authorization Code", vbCritical, "Access Denied"
       'DoCmd.OpenForm "FAMainMEnu"
   Else
        DoCmd.OpenForm "AcctMenu"
   Me.AcctPW.SetFocus
        DoCmd.close acForm, "AcctAuth"
   End If

End Sub

Private Sub CloseMe_Click()
  DoCmd.OpenForm "FAMainMenu"
  DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub Form_Load()
 Me.AcctPW = Null
End Sub
