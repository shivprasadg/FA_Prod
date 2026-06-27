Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz As String


Private Sub AssetID_DblClick(Cancel As Integer)
Call OpenUnitsForm(Me.AssetID)
'DoCmd.OpenForm "UnitsForm", , , "UNITID=" & Me.AssetID
Me.Status.SetFocus
End Sub

Private Sub Buyer_DblClick(Cancel As Integer)
If Me.Buyer = "None" Then Exit Sub
   Call OpenBuyersFrom(Me.BuyerId, 0)
End Sub

Private Sub CompanyName_DblClick(Cancel As Integer)
    DoCmd.close acForm, "ClientsView"
    DoCmd.OpenForm "ClientsView", , , , , , Me.clientID
    Me.Status.SetFocus
End Sub

Private Sub ExcelExport_Click()
FormatExcel "RemarketWrongStatusUnits", "C:\Faas\" & FAUserName & "\Temp\WrongStatus.xlsx", "wrong", , , , "Remaketing Wrong Status Units", True
End Sub

Private Sub Form_Current()
On Error GoTo FixMe
Me.txtCurrent = Me.AssetID
FixMe:
End Sub
