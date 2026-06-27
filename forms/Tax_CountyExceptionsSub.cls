Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz       As String
Dim CheckUpdate As String
Dim CurrentRate As Double



Private Sub Form_Current()
    On Error GoTo FixMe
    Me.txtCurrent = Me.ID
    Me.Parent.TaxTypeId = Me.TaxTypeId
FixMe:
End Sub

Private Sub TaxNumber_AfterUpdate()
    CheckUpdate = AskAboutTaxChange(MakePercent(Me.TaxNumber), MakePercent(CStr(CurrentRate)), Me.TaxNumber)
    If CheckUpdate = "Undo" Then Me.Undo
    If CheckUpdate = "Correction" Then Me.TaxNumber = MakePercent(Me.TaxNumber)
    If CheckUpdate = "AddTax" Then
        DoCmd.close acForm, "Tax_Add"
        DoCmd.OpenForm "Tax_Add", acNormal, , , , , "State|" & Me.TaxTypeId & "|" & Me.Parent.StateID & "|" & Me.TaxNumber
        Me.Undo
    End If
End Sub

Public Sub TaxNumber_GotFocus()
    CurrentRate = Me.TaxNumber
End Sub
