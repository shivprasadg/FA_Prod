Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub CompanyName_Click()
MessageUser "Open", "Opening Client's Form", "just a moment.."
DoCmd.OpenForm "ClientsView", , , , , , Me.ClientID
MessageUser "Close"
End Sub

Private Sub MasterLease_Click()
If Me.MLID & 0 <> 0 Then
MessageUser "Open", "Opening Master Lease Form", "just a moment.."
DoCmd.OpenForm "MstrLease", , , "MLID =" & Me.MLID
MessageUser "Close"
End If
End Sub
