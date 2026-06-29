Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim Askit, PVendor, PType As String



Private Sub Form_Current()
On Error Resume Next
Me.txtCurrent = Me.[VIN]
Me.Parent.AssetID = Me.AssetID
'Me.Parent.RMKTExpenseId = Me.RMKTExpenseId
'Me.Parent.VendorID = Me.VendorID
On Error GoTo 0
End Sub

Private Sub AssetID_DblClick(Cancel As Integer)
   Call OpenUnitsForm(Me.AssetID)
   Me.unitvin.SetFocus
   End Sub

Private Sub UnitVIN_DblClick(Cancel As Integer)
    Call OpenUnitsForm(Me.AssetID)
    Me.AssetID.SetFocus
End Sub
