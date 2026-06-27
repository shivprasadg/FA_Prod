Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub DeleteMe_Click()
    Dim Response, SQLz As String
    Response = MsgBox("You're about to Delete Unit# " & Me.UnitUnitNum & "'s " & vbNewLine & Me.OptionTypeDesc & " option entry , " & vbNewLine & "Are you sure?", vbQuestion + vbYesNo + vbCritical, "What? Delete!")
    If Response = vbNo Then Exit Sub
     On Error GoTo FixMe
    SQLz = "Delete From EXITOptions_WorkTable WHERE UnitID = " & Me.UnitID & "and OptionTypeID = " & Me.OptionTypeID & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    SQLz = "Delete From UnitEXITOption  WHERE LeaseOptionID = " & Me.LeaseOptionID & ";"
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Me.Form.Requery
    Me.Parent.Units_List.Requery
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub Form_Current()
On Error Resume Next
Me.txtCurrent = Me.UnitID
On Error GoTo 0
End Sub

Private Sub UnitID_DblClick(Cancel As Integer)
  Call OpenUnitsForm(Me.UnitID)
  Forms!Unitsform.UnitsTabControl.Value = 8
End Sub
