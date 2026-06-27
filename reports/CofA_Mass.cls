Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Report_Load()
If Forms!CofASignerPopUp.HideSignatureDate = -1 Then Me.SignDateLabel.Visible = False Else Me.SignDateLabel.Visible = True
If Forms!CofASignerPopUp.HideAcceptedDate = -1 Then Me.AcceptedDateLabel.Visible = False Else Me.AcceptedDateLabel.Visible = True
If Forms!CofASignerPopUp.HideAcceptedDate = -1 Then Me.UnitAcceptDate.Visible = False Else Me.UnitAcceptDate.Visible = True
If IsNull(Me.PLBLabel) Then Me.PLBMilesHours.Visible = False Else Me.PLBMilesHours.Visible = True
End Sub
