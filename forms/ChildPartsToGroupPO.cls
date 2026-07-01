Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz As String

Private Sub CloseMe_Click()
    DoCmd.close acForm, (Me.Form.Name)
End Sub

Private Sub GetAll_Click()
    For i = 0 To UnitsList.ListCount - 1
        UnitsList.Selected(i) = True
    Next i
End Sub

Private Sub GetAllBack_Click()
    For i = 0 To POChildPartList.ListCount - 1
        POChildPartList.Selected(i) = True
    Next i
End Sub

Private Sub MoveUnits_Click()
    Dim ItemIndex As Variant
    Call DelWorkTable
    For Each ItemIndex In Me.UnitsList.ItemsSelected
        SQLz = "INSERT INTO AddToGroupPO_WorkTable ( ChildPartID ) SELECT " & Me.UnitsList.ItemData(ItemIndex) & " AS ChildPartID;"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
    Next
    On Error GoTo fixme1
    SQLz = "UPDATE ChildParts INNER JOIN AddToGroupPO_WorkTable ON ChildParts.ChildPartID = AddToGroupPO_WorkTable.ChildPartID SET ChildParts.PONumber = '" & Me.ActivePO & "';"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Me.POChildPartList.Requery
    Me.UnitsList.Requery
    Exit Sub
fixme1:
End Sub
Private Sub MoveUnitsBack_Click()
    Dim ItemIndex As Variant
    Call DelWorkTable
    For Each ItemIndex In Me.POChildPartList.ItemsSelected
        SQLz = "INSERT INTO AddToGroupPO_WorkTable ( ChildPartID ) SELECT " & Me.POChildPartList.ItemData(ItemIndex) & " AS ChildPartID;"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
    Next
    On Error GoTo fixme1
    SQLz = "UPDATE ChildParts INNER JOIN AddToGroupPO_WorkTable ON ChildParts.ChildPartID = AddToGroupPO_WorkTable.ChildPartID SET ChildParts.PONumber = Null;"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Me.POChildPartList.Requery
    Me.UnitsList.Requery
    Exit Sub
fixme1:
End Sub
Private Sub DelWorkTable()
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("DELETE * From AddToGroupPO_WorkTable;")
    DoCmd.SetWarnings True
End Sub

Private Sub Form_Load()
If Me.OpenArgs <> "" Then
Dim GroupIDx
Dim PONumx
PONumx = Trim(SplitString(Me.OpenArgs, "|", 1))
GroupIDx = CLng(SplitString(Me.OpenArgs, "|", 0))
Me.ActivePO = PONumx
Me.SelectCompany = DLookup("clientID", "vw_SixKeys", "SGrpID=" & GroupIDx & "")
Call SelectCompany_Click
Me.LeaseList = DLookup("MLID", "vw_SixKeys", "SGrpID=" & GroupIDx & "")
Call Leaselist_Click
Me.ScheduleList = DLookup("SchID", "vw_SixKeys", "SGrpID=" & GroupIDx & "")
Call ScheduleList_Click
Me.GroupList = GroupIDx
Call GroupList_Click
Me.POList.Requery
End If
End Sub
Public Sub SetFormData()
Me.UnitsList.Requery
Me.SearchUnitID.Requery
Me.SearchPartType.Requery
Me.SearchChildMake.Requery
Me.SearchChildClass.Requery
Me.POChildPartList.Requery
Me.POList.Requery
End Sub
Public Sub ClearFormData()
Me.UnitsList.RowSource = ""
Me.SearchUnitID = ""
Me.SearchPartType = ""
Me.SearchChildClass = ""
Me.SearchChildClass = ""
Me.SearchUnitID.Requery
Me.SearchPartType.Requery
Me.SearchChildMake.Requery
Me.SearchChildClass.Requery
Me.POChildPartList.Requery
Me.POList.Requery
End Sub

Private Sub SearchChildMake_Click()
If Me.SearchChildMake <> "" Then
SQLz = "SELECT ChildPartID, UnitID, ChildType, ChildMake, ChildClass, childFinalCost, GroupREF, PONumber " & _
       "FROM Units INNER JOIN ChildParts ON Units.UnitID = ChildParts.UnitREF " & _
       "WHERE GroupREF=[Forms]![ChildPartsToGroupPO]![GroupList] AND PONumber Is Null " & _
       "AND ChildMake='" & Me.SearchChildMake & "' " & _
       "ORDER BY ChildParts.ChildType, ChildParts.ChildMake, ChildParts.ChildClass;"
Me.UnitsList.RowSource = SQLz
End If
End Sub

Private Sub SearchChildClass_Click()
If Me.SearchChildClass <> "" Then
SQLz = "SELECT ChildPartID, UnitID, ChildType, ChildMake, ChildClass, childFinalCost, GroupREF, PONumber " & _
       "FROM Units INNER JOIN ChildParts ON Units.UnitID = ChildParts.UnitREF " & _
       "WHERE GroupREF=[Forms]![ChildPartsToGroupPO]![GroupList] AND PONumber Is Null " & _
       "AND ChildClass='" & Me.SearchChildClass & "' " & _
       "ORDER BY ChildParts.ChildType, ChildParts.ChildMake, ChildParts.ChildClass;"
Me.UnitsList.RowSource = SQLz
End If
End Sub

Private Sub SearchPartType_Click()
If Me.SearchPartType <> "" Then
SQLz = "SELECT ChildPartID, UnitID, ChildType, ChildMake, ChildClass, childFinalCost, GroupREF, PONumber " & _
       "FROM Units INNER JOIN ChildParts ON Units.UnitID = ChildParts.UnitREF " & _
       "WHERE GroupREF=[Forms]![ChildPartsToGroupPO]![GroupList] AND PONumber Is Null " & _
       "AND ChildType='" & Me.SearchPartType & "' " & _
       "ORDER BY ChildParts.ChildType, ChildParts.ChildMake, ChildParts.ChildClass;"
Me.UnitsList.RowSource = SQLz
End If
End Sub

Private Sub SearchUnitID_Click()
If Me.SearchUnitID <> "" Then
SQLz = "SELECT ChildPartID, UnitID, ChildType, ChildMake, ChildClass, childFinalCost, GroupREF, PONumber " & _
       "FROM Units INNER JOIN ChildParts ON Units.UnitID = ChildParts.UnitREF " & _
       "WHERE GroupREF=[Forms]![ChildPartsToGroupPO]![GroupList] AND PONumber Is Null " & _
       "AND UnitID=" & Me.SearchUnitID & " " & _
       "ORDER BY ChildParts.ChildType, ChildParts.ChildMake, ChildParts.ChildClass;"
Me.UnitsList.RowSource = SQLz
End If
End Sub

Public Sub SelectCompany_Click()
    SQLz = "SELECT MLID, MLNo, MLOrig FROM MstrLease WHERE Client=" & Me.SelectCompany & " ORDER By MLNo;"
    Me.LeaseList.RowSource = SQLz
    Call ClearFormData
End Sub
Private Sub Leaselist_Click()
'    SQLz = "SELECT SchID, SchNo FROM Schedule WHERE SchMLRef=" & Me.LeaseList & " ORDER BY SCHID Desc, SCHNo1;"
    Me.ScheduleList.Requery
Call ClearFormData
End Sub
Private Sub ScheduleList_Click()
'    SQLz = "SELECT SGrpID, SGrpNO FROM SchGrp WHERE SGrpSchID=" & Me.ScheduleList & " ORDER BY SGrpNO ;"
    Me.GroupList.Requery
    Me.GroupID = 0
Call ClearFormData
End Sub
Private Sub GroupList_Click()
SQLz = "SELECT ChildPartID, UnitID, ChildType, ChildMake, ChildClass, childFinalCost, GroupREF, ChildParts.PONumber " & _
       "FROM Units INNER JOIN ChildParts ON Units.UnitID = ChildParts.UnitREF " & _
       "WHERE GroupREF=[Forms]![ChildPartsToGroupPO]![GroupList] AND ChildParts.PONumber Is Null " & _
       "ORDER BY ChildParts.ChildType, ChildParts.ChildMake, ChildParts.ChildClass;"
Me.UnitsList.RowSource = SQLz
Me.GroupID = Nz(Me.GroupList, 0)
Me.SearchUnitID = ""
Me.SearchPartType = ""
Me.SearchChildClass = ""
Me.SearchChildClass = ""
Me.SearchUnitID.Requery
Me.SearchPartType.Requery
Me.SearchChildClass.Requery
Me.SearchChildClass.Requery
Me.POChildPartList.Requery
End Sub
