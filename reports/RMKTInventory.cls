Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub ClientShNmd_DblClick(Cancel As Integer)
DoCmd.close acReport, "RMKTInventoryClient"
DoCmd.OpenReport "RMKTInventoryClient", acViewReport, , , , "SELECT * FROM RMKTInventoryFinal WHERE ClientId=" & Me.clientID & ";"
Me.LocationHDR.SetFocus
End Sub
