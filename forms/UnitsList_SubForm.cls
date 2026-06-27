Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz As String


Private Sub ClientCBO_Click()
Call DelWorkTable
Call MakeWorkTable
Me.Form.Requery
End Sub

Private Sub ModelSubmit_Click()
    On Error GoTo FixMe
    SQLz = "UPDATE Units INNER JOIN MakeModelManager_WorkTable ON Units.UnitID = MakeModelManager_WorkTable.UnitID " & _
           "SET Units.UnitModel = '" & Me.ModelCbo & "' ;"
    DoCmd.RunSQL (SQLz)
FixMe:
    Me.Form.Refresh
    Me.Parent.ModelList.Requery
    Me.Parent.NoMatchModelsList.Requery
    Exit Sub
End Sub
Private Sub UpdateLists()
    Me.Parent.ModelList.Requery
    Me.Parent.NoMatchModelsList.Requery
End Sub
Private Sub UnitID_DblClick(Cancel As Integer)
  ' DoCmd.close acForm, "UnitsForm"
   Call OpenUnitsForm(Me.UnitID)
   'DoCmd.OpenForm "UnitsForm", , , "UNITID=" & Me.UnitID
   'Me.MLNo.SetFocus
   End Sub

Private Sub UnitVIN_DblClick(Cancel As Integer)
Call UnitID_DblClick(1)
End Sub
Private Sub MakeWorkTable()
SQLz = "INSERT INTO MakeModelManager_WorkTable ( UnitID ) SELECT UnitID From Units " & _
"WHERE ClientID = " & Me.ClientCbo.Column(0) & " AND UnitMake = '" & Me.Parent.UMake & "' AND UnitModel ='" & Me.Parent.UModel & "';"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
 End Sub
Private Sub DelWorkTable()
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("DELETE * From MakeModelManager_WorkTable;")
    DoCmd.SetWarnings True
End Sub
