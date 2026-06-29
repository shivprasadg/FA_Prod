Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private Sub CloseformBtn_Click()
    DoCmd.close
End Sub

Private Sub CloseMe_Click()
      If Me.OpenArgs Like "*:BC*" Then DoCmd.OpenForm Split(Me.OpenArgs, ":")(0)
      DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub dealGrandTotal_Click()
    Me.dealGrandTotal = Me.dealUnitQty * Me.dealUnitPrice
End Sub

Private Sub dealUnitPrice_AfterUpdate()
    Me.dealGrandTotal = Me.dealUnitQty * Me.dealUnitPrice
End Sub

Private Sub dealUnitQty_AfterUpdate()
    Me.dealGrandTotal = Me.dealUnitQty * Me.dealUnitPrice
End Sub

Private Sub Form_Close()
    If IsOpen("ClientsView") Then Forms!ClientsView.SetFocus
End Sub

Private Sub Form_GotFocus()
    DoCmd.Maximize
End Sub

Private Sub Form_Load()
If FAUserAccessGroup Like "*ITT*" Or FAUserAccessGroup Like "*TMT*" Then
Me.dealOutcomeDate.Locked = False
End If

'  If IsOpen("ClientsView") Then clientRef = Forms!ClientsView.clientID
End Sub
Private Sub DelDealBtn_Click()
    DoCmd.RunCommand acCmdSelectRecord
    DoCmd.RunCommand acCmdDeleteRecord
End Sub

Private Sub ClientViewFormBtn_Click()
    DoCmd.OpenForm "ClientsView", , , "[clientID]=" & Me.clientREF
End Sub
