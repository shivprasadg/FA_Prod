Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub BuyerName_DblClick(Cancel As Integer)
   If IsNull(Me.BuyerID) Then Exit Sub
   Call OpenBuyersFrom(Me.BuyerID, 0)
End Sub

Private Sub AddBlankUnit_Click()
    If Me.UnitID & "Test" = "Test" Then Exit Sub
    DoCmd.close acForm, "UnitsForm"
    Call OpenUnitsForm(0, Me.GroupRef)
    'DoCmd.OpenForm "UnitsForm", , , "groupref = " & Me.GroupRef, acFormAdd
    Forms!Unitsform.Form.Repaint
    Forms!Unitsform!GroupRef.Value = Me.GroupRef
End Sub
Private Sub CloneMe_Click()
    If Me.UnitID & "Test" = "Test" Then Exit Sub
    DoCmd.close acForm, "UnitsForm"
    Call OpenUnitsForm(Me.UnitID)
    'DoCmd.OpenForm "UnitsForm", , , "UnitID =" & Me.UnitID
    Forms!Unitsform.NoUnits = 1
    Form_UnitsForm.AddUnits_Click
    DoCmd.close acForm, "UnitsForm"
    'DoCmd.OpenForm "UnitsForm", , , "UnitID =" & DMax("UnitID", "Units", "GroupRef=" & Me.GroupRef)
End Sub

Private Sub Form_Load()
    If (FAUserAccess > 11 Or FAUserAccessGroup Like "*OPS*") Or FAUserAcess > 90 Then
        Me.AddBlankUnit.Visible = True
        Me.CloneMe.Visible = True
    Else
        Me.AddBlankUnit.Visible = False
        Me.CloneMe.Visible = False
    End If
End Sub
