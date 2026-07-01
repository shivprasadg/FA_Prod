Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Purchasing_Click()
Call StartSignOff(Me.GroupID, "SignOff1")
Me.Form.Requery
End Sub

Private Sub InvIn_Click()
Call StartSignOff(Me.GroupID, "SignOff2")
Me.Form.Requery
End Sub

Private Sub MSOValid_Click()
Call StartSignOff(Me.GroupID, "SignOff3")
Me.Form.Requery
End Sub

Private Sub Operations_Click()
Call StartSignOff(Me.GroupID, "SignOff4")
Me.Form.Requery
End Sub

Private Sub TMT_Click()
Call StartSignOff(Me.GroupID, "SignOff5")
Me.Form.Requery
End Sub

Private Sub Docs_Click()
Call StartSignOff(Me.GroupID, "SignOff6")
Me.Form.Requery
End Sub
Private Sub Syndication_Click()
Call StartSignOff(Me.GroupID, "SignOff7")
Me.Form.Requery
End Sub
