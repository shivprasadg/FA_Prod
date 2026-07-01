Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz       As String
Dim xSql       As String
Dim RST        As DAO.Recordset
Public Sub ClearBoxes()
    Dim Ccrl   As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Medit" Then
            Me(Ccrl.Name).Value = ""
        End If
    Next Ccrl
End Sub

Private Sub ClearClientSearch_Click()
    Me.SearchClient = ""
    Call SearchClient_Change
End Sub

Private Sub CloseMe_Click()
    DoCmd.OpenForm "FAMainMenu"
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub Form_Load()
    Me.WorkingWith = "Please select a client..."
    SQLz = "SELECT clientID, ClientShortName, CompanyName, Count(AssetID) AS [Unit Count] " & _
        "From vw_SixKeys " & _
        "GROUP BY clientID, ClientShortName, CompanyName " & _
        "ORDER BY CompanyName;"
    Me.ClientList.RowSource = SQLz
    Call ClearAuditLogForm
    Call ClearSearch("All")
    Call ClearBoxes
    Me.LeaseList.RowSource = ""
    Me.ScheduleList.RowSource = ""
    Me.GroupList.RowSource = ""
    Me.Units_List.RowSource = ""
    Me.UnitCount.Visible = False
End Sub

Private Sub GetUnitCounts()
    Dim UnitX, UnitZ As String
    If ([Units_List].[ListCount] - 1) = 1 Then UnitX = "Unit " & ChrW(9658) Else UnitX = "Units " & ChrW(9658)
    If [Units_List].[ItemsSelected].[Count] = 1 Then UnitZ = "Unit " & ChrW(9658) Else UnitZ = "Units " & ChrW(9658)
End Sub
Public Sub ClientList_Click()
    SQLz = "SELECT MLID, MLNo, MLOrig, ClientID From vw_SixKeys " & _
        "GROUP BY MLID, MLNo, MLOrig, ClientID Having ClientID =" & Me.ClientList & " ORDER BY MLOrig;"
    Me.LeaseList.RowSource = SQLz
    Call ClearAuditLogForm
    Me.ScheduleList.RowSource = ""
    Me.GroupList.RowSource = ""
    Me.Units_List.RowSource = ""
    Me.SchID = Null
    Me.GroupID = Null
    Me.UnitID = Null
'    Call Me.AuditManagerLog.Form.HideDetails
    Me.WorkingWith = Me.ClientList.Column(1)
    'Call GetUnitCounts
End Sub
Public Sub Leaselist_Click()
    SQLz = "SELECT SchID, Schedule, MLID, Count(AssetID) AS Unt, ClientID, MLNo, MLOrig " & _
        "FROM vw_SixKeys " & _
        "GROUP BY SchID, Schedule, MLID, ClientID, MLNo, MLOrig HAVING MLID=" & Me.LeaseList & " " & _
        "ORDER BY CDbl(NumInString(Left(Schedule,3))), Schedule;"
    ' Sqlz = "SELECT * From ReMarketing_MLSchGrpUnits WHERE MLID=" & Me.LeaseList & ";"
    Me.ScheduleList.RowSource = SQLz
    Call ClearAuditLogForm
    Me.SchID = Null
    Me.GroupID = Null
    Me.UnitID = Null
    Call Me.AuditManagerLog.Form.HideDetails
    Me.WorkingWith = Me.ClientList.Column(1) & " MLNo: " & LeaseList.Column(1)
    Me.GroupList.RowSource = ""
    Me.Units_List.RowSource = ""
   ' Call GetUnitCounts
End Sub

Public Sub ScheduleList_Click()
    Me.SchID = Me.ScheduleList
    Me.GroupID = Null
    Me.UnitID = Null
    Me.Units_List.RowSource = ""
    Call Me.AuditManagerLog.Form.HideDetails
    Me.WorkingWith = Me.ClientList.Column(1) & " MLNo: " & LeaseList.Column(1) & " Sch: " & Me.ScheduleList.Column(1)
    SQLz = "SELECT SchID, sGrpID As GroupID, AssetID, UnitNum, VIN, [Status], Make, SubType, MYear, UnitGroup " & _
        "FROM vw_SixKeys " & _
        "WHERE SchID=" & Me.ScheduleList & " ORDER BY AssetID;"
    Me.Units_List.RowSource = SQLz

    SQLz = "SELECT SgrpID, UnitGroup, Count(AssetID) AS Unt, SchID " & _
        "FROM vw_SixKeys " & _
        "GROUP BY SgrpID, UnitGroup,SchID HAVING SchID=" & Me.ScheduleList & " ORDER BY UnitGroup;"
    Me.GroupList.RowSource = SQLz
    SQLz = ""
    SQLz = "SELECT identity, ActionDate, ActionDesc, Nz([OldValue],0) AS Prev, Nz([NewValue],0) AS New, " & _
        "SysUser AS [User], FieldName, GroupID, ATFGroupSpecID, SchID, FieldGroup " & _
        "FROM vw_FAID_RevGroupUnion " & _
        "WHERE SchID= " & Me.ScheduleList.Column(0) & " " & _
        "ORDER BY FieldName, ActionDate Desc;"
    Me.AuditManagerLog.Form.RecordSource = SQLz
    SQLz = "SELECT FieldName " & _
        "FROM vw_FAID_RevGroupUnion " & _
        "WHERE SchID= " & Me.ScheduleList.Column(0) & " " & _
        "Group By FieldName " & _
        "ORDER BY FieldName;"
    Me.AuditManagerLog!Search_FieldName.RowSource = SQLz
    SQLz = "SELECT SysUser " & _
        "FROM vw_FAID_RevGroupUnion " & _
        "WHERE SchID= " & Me.ScheduleList.Column(0) & " " & _
        "Group By SysUser " & _
        "ORDER BY SysUser;"
    Me.AuditManagerLog!Search_SysUser.RowSource = SQLz
    Call ClearSearch("All")
    'Call GetUnitCounts
End Sub
Public Sub GroupList_Click()
    Me.UnitID = Null
    Call Me.AuditManagerLog.Form.HideDetails
    Me.GroupID = Me.GroupList.Column(0)
    SQLz = "SELECT SchID, sGrpID As GroupID, AssetID, UnitNum, VIN, [Status], Make, SubType, MYear, UnitGroup " & _
        "FROM vw_SixKeys " & _
        "WHERE sGrpID=" & Me.GroupList & " ORDER BY AssetID;"
    Me.Units_List.RowSource = SQLz

    SQLz = "SELECT identity, ActionDate, ActionDesc, Nz([OldValue],0) AS Prev, Nz([NewValue],0) AS New, " & _
        "SysUser AS [User], FieldName, GroupID, ATFGroupSpecID, SchID, FieldGroup " & _
        "FROM vw_FAID_RevGroupUnion " & _
        "WHERE GroupID= " & Me.GroupList.Column(0) & " " & _
        "ORDER BY FieldName, ActionDate Desc;"
    Me.AuditManagerLog.Form.RecordSource = SQLz
    SQLz = "SELECT FieldName " & _
        "FROM vw_FAID_RevGroupUnion " & _
        "WHERE GroupID= " & Me.GroupList.Column(0) & " " & _
        "Group By FieldName " & _
        "ORDER BY FieldName;"
    Me.AuditManagerLog!Search_FieldName.RowSource = SQLz
    SQLz = "SELECT SysUser " & _
        "FROM vw_FAID_RevGroupUnion " & _
        "WHERE GroupID= " & Me.GroupList.Column(0) & " " & _
        "Group By SysUser " & _
        "ORDER BY SysUser;"
    Me.AuditManagerLog!Search_SysUser.RowSource = SQLz
    Me.WorkingWith = Me.ClientList.Column(1) & " MLNo: " & LeaseList.Column(1) & " Sch: " & Me.ScheduleList.Column(1) & " Grp: " & GroupList.Column(1)
    Call ClearSearch("All")
    'Call GetUnitCounts
End Sub
Private Sub ClearAuditLogForm()
    SQLz = "SELECT identity, ActionDate, ActionDesc, Nz([OldValue],0) AS Prev, Nz([NewValue],0) AS New, " & _
        "SysUser AS [User], FieldName, GroupID, ATFGroupSpecID, SchID, FieldGroup " & _
        "FROM vw_FAID_RevGroupUnion " & _
        "WHERE SchID= 0 "
    Me.AuditManagerLog.Form.RecordSource = SQLz
End Sub
Public Function ClearSearch(NotMe As String)
    If NotMe <> "SearchAsset" Then Me.SearchAsset = ""
    If NotMe <> "SearchUnit" Then Me.SearchUnit = ""
    If NotMe <> "SearchVin" Then Me.SearchVin = ""
    If NotMe <> "SearchStatus" Then Me.SearchStatus = ""
    If NotMe <> "SearchMake" Then Me.SearchMake = ""
    If NotMe <> "SearchType" Then Me.SearchType = ""
End Function
Private Sub SearchAsset_Change()
    Call ClearSearch("SearchAsset")
    SQLz = Me.Units_List.RowSource
    SQLz = SplitString(SQLz, " ORDER BY", 0)
    SQLz = SplitString(SQLz, " AND", 0)
    SQLz = SQLz & " AND AssetId Like '*" & Me.SearchAsset.Text & "*' ORDER BY AssetID;"
    Me.Units_List.RowSource = SQLz
  '  Call GetUnitCounts
End Sub

Private Sub SearchClient_Change()
    SQLz = "SELECT clientID, ClientShortName, CompanyName, Count(AssetID) AS [Unit Count] " & _
        "From vw_SixKeys WHERE ClientShortName Like '*" & Me.SearchClient.Text & "*' Or CompanyName Like '*" & Me.SearchClient.Text & "*' " & _
        "GROUP BY clientID, ClientShortName, CompanyName " & _
        "ORDER BY CompanyName;"
    Me.ClientList.RowSource = SQLz
End Sub

Private Sub SearchUnit_Change()
    Call ClearSearch("SearchUnit")
    SQLz = Me.Units_List.RowSource
    SQLz = SplitString(SQLz, " ORDER BY", 0)
    SQLz = SplitString(SQLz, " AND", 0)
    SQLz = SQLz & " AND UnitNum Like '*" & Me.SearchUnit.Text & "*' ORDER BY UnitNum;"
    Me.Units_List.RowSource = SQLz
    'Call GetUnitCounts
End Sub
Private Sub SearchVIN_Change()
    Call ClearSearch("SearchVin")
    SQLz = Me.Units_List.RowSource
    SQLz = SplitString(SQLz, " ORDER BY", 0)
    SQLz = SplitString(SQLz, " AND", 0)
    SQLz = SQLz & " AND VIN Like '*" & Me.SearchVin.Text & "*' ORDER BY VIN;"
    Me.Units_List.RowSource = SQLz
    'Call GetUnitCounts
End Sub
Private Sub SearchStatus_Change()
    Call ClearSearch("SearchStatus")
    SQLz = Me.Units_List.RowSource
    SQLz = SplitString(SQLz, " ORDER BY", 0)
    SQLz = SplitString(SQLz, " AND", 0)
    SQLz = SQLz & " AND [Status] Like '*" & Me.SearchStatus.Text & "*' ORDER BY AssetID;"
    Me.Units_List.RowSource = SQLz
   ' Call GetUnitCounts
End Sub
Private Sub SearchMake_Change()
    Call ClearSearch("SearchMake")
    SQLz = Me.Units_List.RowSource
    SQLz = SplitString(SQLz, " ORDER BY", 0)
    SQLz = SplitString(SQLz, " AND", 0)
    SQLz = SQLz & " AND Make Like '*" & Me.SearchMake.Text & "*' ORDER BY AssetID;"
    Me.Units_List.RowSource = SQLz
    'Call GetUnitCounts
End Sub
Private Sub SearchType_Change()
    Call ClearSearch("SearchType")
    SQLz = Me.Units_List.RowSource
    SQLz = SplitString(SQLz, " ORDER BY", 0)
    SQLz = SplitString(SQLz, " AND", 0)
    SQLz = SQLz & " AND SubType Like '*" & Me.SearchType.Text & "*' ORDER BY AssetID;"
    Me.Units_List.RowSource = SQLz
    'Call GetUnitCounts
End Sub
Private Sub ClearLabel_Click()
    Me.SearchVin.SetFocus
    Call SearchVIN_Change
End Sub

Public Sub Units_List_Click()
    Me.UnitID = Me.Units_List.Column(2)
    Call Me.AuditManagerLog.Form.HideDetails
    SQLz = "SELECT AssetID as [identity], ActionDate, [Action] as ActionDesc,"
    SQLz = SQLz & "Nz(OldLocation,[OldValue]) AS Prev, NZ(Newlocation,[NewValue]) AS New, SysUser AS [User], FieldLabel as FieldName, GroupID "
    SQLz = SQLz & "FROM vw_AuditUnits "
    SQLz = SQLz & "WHERE AssetID= " & Me.UnitID & " "
    SQLz = SQLz & "ORDER BY FieldLabel, ActionDate Desc;"
    Me.AuditManagerLog.Form.RecordSource = SQLz

    SQLz = "SELECT FieldLabel as  FieldName " & _
        "FROM vw_AuditUnits " & _
        "WHERE AssetID= " & Me.UnitID & " " & _
        "Group By FieldLabel " & _
        "ORDER BY FieldLabel;"
    Me.AuditManagerLog!Search_FieldName.RowSource = SQLz
    SQLz = "SELECT SysUser " & _
        "FROM vw_AuditUnits " & _
        "WHERE AssetID= " & Me.UnitID & " " & _
        "Group By SysUser " & _
        "ORDER BY SysUser;"
    Me.AuditManagerLog!Search_SysUser.RowSource = SQLz
    Me.WorkingWith = Me.ClientList.Column(1) & " MLNo: " & LeaseList.Column(1) & " Sch: " & Me.ScheduleList.Column(1) & " Grp: " & GroupList.Column(1) & " AssetID: " & Units_List.Column(2)
    'Call GetUnitCounts
End Sub

Private Sub Units_List_DblClick(Cancel As Integer)
    On Error GoTo FixMe

FixMe:
End Sub
