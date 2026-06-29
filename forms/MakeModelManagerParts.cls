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
'Call DelWorkTable
'Me.UnitPartsList_SubForm.Form.Requery
End Sub

Private Sub MakeList_Click()
    SQLz = "SELECT ChildPartModelID, ChildPartModel AS Models, PartType, ChildPartMakeID " & _
           "From LookUp_ChildPartModel " & _
           "WHERE ChildPartMakeID =" & Me.MakeList & " Order By ChildPartModel;"
    Me.ModelList.RowSource = SQLz
    'SQLz = "SELECT GroupMake As [UnMatched Model], Leased As Lease, Tracked as Track, UnitMake " & _
           "From LookUp_ChildPartModel WHERE ChildPartMakeId=" & Me.MakeList.Column(1) & " " & _
           "ORDER BY Units_MakesModel_NotListed.UnitModel;"
    Call DelWorkTable
    Me.NoMatchModelsList.RowSource = SQLz
    Me.AddBox.Visible = False
    Me.AddMissingModel.Visible = False
    Me.NoMatchModelsList = 0
    Me.Make = Me.MakeList.Column(1)
    Me.Model = Null
    Me.UnitPartsList_SubForm.Form.Requery
   ' Me.UnitPartsList_SubForm.Form.MakeCbo = ""
    Me.UnitPartsList_SubForm.Form.ModelCbo = ""
End Sub

Private Sub ModelList_Click()
    Me.Model = Me.ModelList.Column(1)
    Call MessageUser("Open", "Getting '" & Me.Model & "' details", "just a moment...", False)
    Call DelWorkTable
    Call MakeWorkTable
    'Me.UnitPartsList_SubForm.Form.Requery
    'Me.UnitPartsList_SubForm!ClientCbo.Requery
Call MessageUser("Close")
End Sub
Private Sub MakeWorkTable()
    SQLz = "INSERT INTO MakeModelManagerParts_WorkTable ( UnitID, ChildPartID ) SELECT UnitREF, ChildPartID From ChildParts Where ChildMake='" & Me.Make & "' AND ChildClass='" & Me.Model & "'"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
End Sub
Private Sub DelWorkTable()
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("DELETE * From MakeModelManagerParts_WorkTable;")
    DoCmd.SetWarnings True
End Sub

'Private Sub NoMatchModelsList_Click()
   ' Me.AddBox.Visible = True
   ' Me.AddMissingModel.Visible = True
   ' Me.UMake = Me.NoMatchModelsList.Column(3)
   ' Me.UModel = Me.NoMatchModelsList.Column(0)
   ' Me.AddMissingModel.caption = "Add " & Me.NoMatchModelsList & vbNewLine & " to " & Me.MakeList.Column(1) & "'s" & vbNewLine & "models list"
   ' Call MessageUser("Open", "Getting '" & Me.UModel & "' details", "just a moment...", False)
   ' Call DelWorkTable
   ' Call MakeWorkTable
   ' Me.UnitPartsList_SubForm.Form.Requery
   ' Me.UnitPartsList_SubForm!ClientCbo.Requery
   ' Call MessageUser("Close")
'End Sub
