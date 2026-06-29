Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub EditBtn_Click()
    DoCmd.close acForm, "Deals_view"
    DoCmd.OpenForm "Deals_view", , , "[dealID]=" & Me![dealID]
End Sub

Private Sub Form_Close()
   On Error GoTo Form_Close_Error
   On Error GoTo 0
   Exit Sub
Form_Close_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Close of VBA Document Form_proposalsSummary_sub"
End Sub

Public Sub SetDeals()
Me.Form.RecordSource = "Select * From Deals Where ClientID=" & Me.Parent.ClientID & ";"
Me.Form.Visible = True
End Sub

Private Sub LOVolumePct_Change()
    Me.LOTotal.Requery
End Sub

Private Sub RefreshBtn_Click()
   Me.Form.Refresh
End Sub
