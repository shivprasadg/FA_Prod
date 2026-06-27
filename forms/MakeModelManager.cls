Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz       As String

Private Sub AddMissingModel_Click()
    SQLz = "INSERT INTO LookUp_UnitModels ( UnitMake, UnitModel, MakeID ) " & _
           "SELECT '" & Me.MakeList.Column(1) & "', '" & Me.NoMatchModelsList.Column(0) & "', " & Me.MakeList.Column(0) & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Me.MakeList.SetFocus
    Call MakeList_Click
End Sub

Private Sub CloseMe_Click()
    DoCmd.close acForm, (Me.Form.Name)
End Sub

Private Sub Form_Load()
Call DelWorkTable
Me.UnitsList_SubForm.Form.Requery
End Sub

Private Sub MakeList_Click()
    Me.ModelList.Requery
    SQLz = "SELECT UnitModel As [UnMatched Model], Leased As Lease, Tracked as Track, UnitMake " & _
           "From Units_MakesModel_NotListed WHERE UnitMake = '" & Me.MakeList.Column(1) & "' " & _
           "ORDER BY Units_MakesModel_NotListed.UnitModel;"
    Call DelWorkTable
    Me.NoMatchModelsList.RowSource = SQLz
    Me.AddBox.Visible = False
    Me.AddMissingModel.Visible = False
    Me.NoMatchModelsList = 0
    Me.UMake = Me.MakeList.Column(1)
    Me.UModel = Null
    Me.UnitsList_SubForm.Form.Requery
    Me.UnitsList_SubForm.Form.MakeCbo = ""
    Me.UnitsList_SubForm.Form.ModelCbo = ""
End Sub

Private Sub ModelList_Click()
Me.UModel = Me.ModelList.Column(1)
    Call MessageUser("Open", "Getting '" & Me.UModel & "' details", "just a moment...", False)
    Call DelWorkTable
    Call MakeWorkTable
    Me.UnitsList_SubForm.Form.Requery
    Me.UnitsList_SubForm!ClientCbo.Requery
Call MessageUser("Close")
End Sub
Private Sub MakeWorkTable()
    SQLz = "INSERT INTO MakeModelManager_WorkTable ( UnitID ) SELECT UnitID From Units " & _
           "WHERE UnitMake = '" & Me.UMake & "' AND UnitModel ='" & Me.UModel & "';"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
End Sub
Private Sub DelWorkTable()
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("DELETE * From MakeModelManager_WorkTable;")
    DoCmd.SetWarnings True
End Sub

Private Sub NoMatchModelsList_Click()
    Me.AddBox.Visible = True
    Me.AddMissingModel.Visible = True
    Me.UMake = Me.NoMatchModelsList.Column(3)
    Me.UModel = Me.NoMatchModelsList.Column(0)
    Me.AddMissingModel.caption = "Add " & Me.NoMatchModelsList & vbNewLine & " to " & Me.MakeList.Column(1) & "'s" & vbNewLine & "models list"
    Call MessageUser("Open", "Getting '" & Me.UModel & "' details", "just a moment...", False)
    Call DelWorkTable
    Call MakeWorkTable
    Me.UnitsList_SubForm.Form.Requery
    Me.UnitsList_SubForm!ClientCbo.Requery
    Call MessageUser("Close")
End Sub
