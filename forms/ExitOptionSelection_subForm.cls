Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub ExchangedByUnitId_DblClick(Cancel As Integer)
DoCmd.close acForm, "UnitSnapShot"
DoCmd.OpenForm "UnitSnapShot", acNormal, , "UnitID = " & Me.ExchangedByUnitId & ""
End Sub

Private Sub OptionSelected_Click()
If Me.OptionSelected = "Undecided" Then Me.OptionTerm = 0
If Me.OptionSelected = "Lease end" Then Me.OptionTerm = 0
If Me.OptionSelected = "M-T-M" Or Me.OptionSelected Like "*change*" Then
 If IsNull(Me.OptionTerm) Or OptionTerm < 3 Then Me.OptionTerm = 3
End If

End Sub
