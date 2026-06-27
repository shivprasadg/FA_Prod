Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub UnitID_DblClick(Cancel As Integer)
  ' DoCmd.close acForm, "UnitsForm"
   Call OpenUnitsForm(Me.UnitID)
   'DoCmd.OpenForm "UnitsForm", , , "UNITID=" & Me.UnitID
   Me.UnitVIN.SetFocus
   End Sub

Private Sub UnitVIN_DblClick(Cancel As Integer)
Call UnitID_DblClick(1)
End Sub
