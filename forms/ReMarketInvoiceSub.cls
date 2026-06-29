Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub AssetID_DblClick(Cancel As Integer)
Call OpenUnitsForm(Me.AssetID)
'DoCmd.OpenForm "UnitsForm", , , "UnitID=" & Me.AssetID & ""
Me.unitvin.SetFocus

End Sub

Private Sub RmktSaleDt_AfterUpdate()
If IsDate(Me.RmktSaleDt) Then Me.UnitStatus = "S" Else Me.UnitStatus = "R"
End Sub

Private Sub UNSaleUnit_Click()
   Call UnSellAUnit(Me.AssetID, Me.InvoiceRequestID)
   Me.Refresh
End Sub
