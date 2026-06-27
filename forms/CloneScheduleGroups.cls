Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim i As Integer
Dim SQLz As String
Dim OldSchNo As String
Dim OldSchDec As String
Dim Askit, Msgx As String
Dim NewSchIDi As Integer
Dim rs As DAO.Recordset

Private Sub CloneSchNGroup_Click()
    If Me.NewSchNo & "TestT" = "TesT" Then
       MsgBox "You need to Re-Enter the New Schedule Num", vbCritical, "You can trust a Monkey!"
       Exit Sub
    End If
    If Me.Dirty Then Me.Dirty = False
    Me.Form.Requery
    Dim nogroups, nogroupnum, nogroupdesc As Integer
    nogroups = DCount("ID", "CloneSchGrp_WorkTable", "CloneMe = -1")
    nogroupnum = DCount("ID", "CloneSchGrp_WorkTable", "CloneMe = -1 and nz(NewGroupNo,'')=''")
    nogroupdesc = DCount("ID", "CloneSchGrp_WorkTable", "CloneMe = -1 and nz(NewGroupDesc,'')=''")

    Msgx = ""
    If Me.NewSchNotoUse & "test" = "test" Then Msgx = Msgx & "    The New SchNo is missing" & vbNewLine
    If Me.NewSchDesctoUse & "test" = "test" Then Msgx = Msgx & "    The New Sch Desc is missing" & vbNewLine
    If nogroups = 0 Then Msgx = Msgx & "    No Groups have been selected to Clone" & vbNewLine
    If nogroupnum > 0 Then Msgx = Msgx & "    " & nogroupnum & " New Group Num(s) are missing" & vbNewLine
    If nogroupdesc > 0 Then Msgx = Msgx & "    " & nogroupdesc & " New Group Descp(s) are missing" & vbNewLine
    If Msgx <> "" Then
        Msgx = "Please correct the the following:" & vbNewLine & Msgx & vbNewLine
        Msgx = Msgx & "The Schedule and the Group have NOT been cloned..."
        MsgBox Msgx, vbCritical, AtlaasVersion
        Exit Sub
    End If
    Askit = MsgBox("You are about to Clone 1 Schedule and (" & DCount("ID", "CloneSchGrp_WorkTable", "CloneMe=-1") & ") Group(s)" & vbNewLine & vbNewLine & _
                   "Are you Sure?", vbQuestion + vbYesNo + vbCritical, "Run Clone-Set")
    If Askit <> 6 Then Exit Sub
    SQLz = "SELECT * FROM CloneSchGrp_WorkTable WHERE CloneMe = -1 "
    Set rs = CurrentDb.OpenRecordset(SQLz)
    rs.MoveLast
    rs.MoveFirst
    NewSchIDi = 0
    Call CloneScheduleGroup(rs!SourceSchID, Me.SourceGroupID, rs!NewSchNo, rs!NewGroupNo, rs!NewSchDesc, rs!NewGroupDesc, 0, 0, 1)
    NewSchIDi = DMax("SchID", "Schedule", "SchMLRef = " & DMax("MLID", "vw_SixKeys", "SchID=" & Me.SchID & "") & "")
    Me.NewSchIDtoUse = NewSchIDi
    For i = 1 To rs.RecordCount
        Call CloneScheduleGroup(rs!SourceSchID, rs!SourceGroupID, rs!NewSchNo, rs!NewGroupNo, rs!NewSchDesc, rs!NewGroupDesc, 1, NewSchIDi, 0)
        rs.MoveNext
    Next i

    rs.MoveFirst
    'Copy ATFChanges
    If DCount("SchID", "ATF_ChangeReasons", "SchID=" & rs!SourceSchID & "") > 0 Then
        SQLz = " INSERT INTO ATF_ChangeReasons " & _
            "(SchID, [Action],Reason,ChangeDate,FieldName,SchNo,GroupNo,FieldGroup,OldValue,NewValue,TableName " & _
            ",RevID,GroupID,UnitID,SysUser,ATFID,ATFUnitSpecID,ATFGroupSpecID,ATFSchSpecID,ReviewedBy,ReviewedDate) " & _
            "SELECT " & Me.NewSchIDtoUse & ", [Action],Reason,ChangeDate,FieldName,'" & rs!NewSchNo & "',GroupNo,FieldGroup,OldValue,NewValue,TableName " & _
            ",RevID,GroupID,UnitID,SysUser,ATFID,ATFUnitSpecID,ATFGroupSpecID,ATFSchSpecID,ReviewedBy,ReviewedDate " & _
            "From ATF_ChangeReasons Where SchID=" & rs!SourceSchID & ";"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
    End If


    Askit = MsgBox("Your Request: Clone 1 Schedule with " & rs.RecordCount & " Group(s), has completed" & _
                   vbNewLine & "Do you want to move units?" & vbNewLine & vbNewLine & _
                   "(this form will close afterwards)", vbQuestion + vbYesNo + vbCritical, "Cloning done")
    If Askit = 6 Then
        DoCmd.OpenForm "MoveUnit"
        Forms!MoveUnit.Form.SelectCompany = DMax("ClientID", "vw_SixKeys", "SchID=" & rs!SourceSchID & "")
        Form_MoveUnit.SelectCompany_Click
        Forms!MoveUnit.LeaseList = DMax("MLID", "vw_SixKeys", "SchID=" & rs!SourceSchID & "")
        Form_MoveUnit.Leaselist_Click
        Forms!MoveUnit.ScheduleList = rs!SourceSchID
        Form_MoveUnit.ScheduleList_Click
        Forms!MoveUnit.GroupList = rs!SourceGroupID
        Form_MoveUnit.GroupList_Click
        Forms!MoveUnit.NewSchedule = Me.NewSchIDtoUse
        Form_MoveUnit.NewSchedule_Click
        Forms!MoveUnit.NewGroupList = DMin("SGrpId", "SchGrp", "sGrpSchID=" & Me.NewSchIDtoUse & "")
        Form_MoveUnit.NewGroupList_Click
    End If
    rs.close
    Set rs = Nothing
    DoCmd.close acForm, Me.Form.Name
End Sub


Private Sub CloseMe_Click()
    Call RePairExchangeIT
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub Form_Load()
    If IsNull(Me.OpenArgs) Then
        Me.SchID = 641
    Else
        Me.SchID = SplitString(Me.OpenArgs, "|", 0)
        SourceGroupID = SplitString(Me.OpenArgs, "|", 1)
    End If
    OldSchNo = DLookup("SchNo", "Schedule", "SchId=" & Me.SchID)
    OldSchDec = DLookup("SchShortDesc", "Schedule", "SchId=" & Me.SchID)
    Me.Form.RecordSource = ""
    Call MakeWorkTable
    Me.Form.RecordSource = "Select * From CloneSchGrp_WorkTable"
    Me.NewSchNotoUse = ""
    Me.NewSchDesctoUse = ""
    Me.NewSchDesctoUse = DMax("NewSchDesc", "CloneSchGrp_WorkTable")
End Sub

Private Sub MakeWorkTable()
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("DELETE FROM CloneSchGrp_WorkTable")
    SQLz = "INSERT INTO CloneSchGrp_WorkTable(SourceSchID,SourceGroupID,SourceSchNo,NewSchDesc,SourceGroupNo,NewGroupDesc) "
    SQLz = SQLz & "SELECT SGrpSchID, SgrpID,'" & OldSchNo & "','" & OldSchDec & "',SGrpNo, SgrpDesc "
    SQLz = SQLz & "FROM SchGrp Where SGrpSchID=" & Me.SchID
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True

End Sub


Private Sub JustCopyGroup_Click()
    Msgx = ""
    If Me.CloneMe = 0 Then Msgx = Msgx & "    Group '" & SourceGroupNo & "' has NOT been selected to Clone" & vbNewLine
    If Me.NewGroupNo & "test" = "test" Then Msgx = Msgx & "    The New Group No is missing" & vbNewLine
    If Me.NewGroupDesc & "test" = "test" Then Msgx = Msgx & "    The New Group Desc is missing" & vbNewLine
    If Msgx <> "" Then
        Msgx = "Please correct the the following:" & vbNewLine & Msgx & vbNewLine
        Msgx = Msgx & "Group '" & SourceGroupNo & "' has NOT been cloned..."
        MsgBox Msgx, vbCritical, AtlaasVersion
        Exit Sub
    End If
    Askit = MsgBox("You are about to Clone Group: '" & SourceGroupNo & "' to NEW Group: '" & Me.NewGroupNo & "'" & vbNewLine & _
                   "Into the Same Schedule " & Me.SourceSchNo & "." & vbNewLine & vbNewLine & _
                   "Are you Sure?", vbQuestion + vbYesNo + vbCritical, "Run Just Clone a Group")
    If Askit <> 6 Then Exit Sub
    Call CloneScheduleGroup(Me.SourceSchID, Me.SourceGroupID, Me.SourceGroupNo, Me.NewGroupNo, Me.NewSchDesc, Me.NewGroupDesc, 1, Me.SourceSchID, 0)
    Askit = MsgBox("Your Request: Clone Group: '" & SourceGroupNo & "' to NEW Group: '" & Me.NewGroupNo & "' has completed" & _
                   vbNewLine & "Do you want to move units?" & vbNewLine & vbNewLine & _
                   "(this form will close afterwards)", vbQuestion + vbYesNo + vbCritical, "Clone Group done")
    If Askit = 6 Then
        DoCmd.OpenForm "MoveUnit"
        Forms!MoveUnit.Form.SelectCompany = DMax("ClientID", "vw_SixKeys", "SchID=" & Me.SourceSchID & "")
        Form_MoveUnit.SelectCompany_Click
        Forms!MoveUnit.LeaseList = DMax("MLID", "vw_SixKeys", "SchID=" & Me.SourceSchID & "")
        Form_MoveUnit.Leaselist_Click
        Forms!MoveUnit.ScheduleList = Me.SourceSchID
        Form_MoveUnit.ScheduleList_Click
        Forms!MoveUnit.GroupList = Me.SourceGroupID
        Form_MoveUnit.GroupList_Click
        Forms!MoveUnit.NewSchedule = Me.SourceSchID
        Form_MoveUnit.NewSchedule_Click
        Forms!MoveUnit.NewGroupList = DMax("SGrpId", "SchGrp", "sGrpSchID=" & Me.SourceSchID & "")
        Form_MoveUnit.NewGroupList_Click
    End If
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub NewSchDesctoUse_AfterUpdate()
    SQLz = "Update CloneSchGrp_WorkTable "
    SQLz = SQLz & "Set NewSchDesc = '" & Me.NewSchDesctoUse & "'"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Me.Form.Requery
End Sub

Private Sub NewSchNotoUse_AfterUpdate()
    If Me.Dirty Then Me.Dirty = False
    Me.NewSchIDtoUse.SetFocus
    SQLz = "Update CloneSchGrp_WorkTable "
    SQLz = SQLz & "Set NewSchNo='" & Me.NewSchNotoUse & "'"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Me.Form.Requery
End Sub

Private Sub NewSchNotoUse_LostFocus()
   If Me.Dirty Then Me.Dirty = False
End Sub
