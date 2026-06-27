Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz, SQLx       As String

Private Sub CloseMe_Click()
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub Form_AfterUpdate()
   Me.Parent.ModMade = -1
End Sub

Private Sub Form_Current()
    Me.txtCurrent = Me.UnitID
End Sub
Private Sub GoHome_Click()
    DoCmd.OpenForm " FAMainMenu"
End Sub

Public Sub ClearBoxes()
    Dim Ccrl   As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Medit" Then
            Me(Ccrl.Name).Value = ""
        End If
    Next Ccrl
End Sub
Private Sub Form_ApplyFilter(Cancel As Integer, ApplyType As Integer)
    Dim zFilter As String
    zFilter = Replace(Replace(Replace(Me.Form.Filter, "[Bulk_Update_SelectedUnits].", ""), "]", ""), "[", "")
    If zFilter <> "" Then Call Me.Parent.SetUnits(zFilter)
    Call Me.Parent.SetOrderBy(Me.OrderBy)
    Me.Form.Filter = ""
End Sub

Private Sub Form_Dirty(Cancel As Integer)
   Me.Parent.ModMade = -1
End Sub

Private Sub Form_Load()
    If FAUserAccess > 50 And FAUserAccessGroup Like "*SYND*" Then
        Me.COAOrgSubmit.Visible = True
        Me.COAOrgSubmit.Visible = True
        Me.COAOrgCbo.Visible = True
        Me.UnitCOABack.Locked = False
    Else
        Me.COAOrgSubmit.Visible = False
        Me.COAOrgSubmit.Visible = False
        Me.COAOrgCbo.Visible = False
        Me.UnitCOABack.Locked = True
    End If
End Sub
Private Sub AcceptedDateSubmit_Click()
    On Error GoTo FixMe
    Dim ExtraFields As String
    If Me.AcceptedDate & "test" <> "test" Then
    ExtraFields = ", Units.UnitDocAccpt = -1, Units.unitstatus = 'A' "
    Call UpdateMassUnits("Units.UnitAcceptDate", Me.AcceptedDate, "", ExtraFields)
    'Call UpdateMassUnits("Units.UnitAcceptDate", Me.AcceptedDate, "WHERE IsNull(Units.UnitAcceptDate)", ExtraFields)
    ModMade = -1
    End If
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub COAOrgSubmit_Click()
    On Error GoTo FixMe
    Dim ExtraFields As String
    If Me.Parent.LeaseList.Column(2) = "FA" Then
        If IsDate(Me.COAOrgCbo) Then
        ExtraFields = ", Units.UnitDocAccpt = -1, Units.unitstatus = 'A' "
        Call UpdateMassUnits("Units.UnitCOABack", Me.COAOrgCbo, "WHERE Units.UnitStatus ='B' or 'A'", ExtraFields)
        End If
    End If
    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub
Private Sub LienReleaseSubmit_Click()
    On Error GoTo FixMe
    If Me.LienReleaseCbo & "test" <> "test" Then
    Call UpdateMassUnits("Units.LienRelease", Me.LienReleaseCbo)
    End If
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub PLBSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitPLB", Me.PLBcbo)
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub UnitCOACopy_AfterUpdate()
    On Error GoTo FixMe
    Dim ExtraFields As String
    If Me.Parent.LeaseList.Column(2) = "FA" Then
        If IsDate(Me.UnitCofACopy) Then
        ExtraFields = ", Units.UnitDocAccpt = -1, Units.unitstatus = 'A' "
        Call UpdateMassUnits("Units.UnitCOABack", Me.UnitCofACopy, "WHERE Units.UnitStatus ='B' or 'A'", ExtraFields)
        End If
    End If
FixMe:
End Sub

Private Sub UnitMSOSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitMSOCopy", Me.UnitMSOCBO)
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub UnitStatus_Click()
    If Me.unitstatus Like "*R*" Or Me.unitstatus Like "*S*" Then
        If DCount("UnitRef", "RMKT", "[UNITRef]=" & Me.UnitID & "") = 0 Then
            SQLz = "INSERT INTO Rmkt (UnitRef) SELECT " & Me.UnitID & " AS UnitId;"
            DoCmd.RunSQL (SQLz)
        End If
    End If
    If Me.unitstatus = "A" Then Call CheckMessage
End Sub


Private Sub UnitStatusSubmit_Click()
    If Me.UnitStatusCbo Like "*R*" Or Me.UnitStatusCbo Like "*S*" Or Me.UnitStatusCbo Like "*D*" Or Me.UnitStatusCbo Like "*E*" Then
        MsgBox "You cannot 'MASS' Update Unit Status on Remarketed Units." & vbNewLine & "You can do it at the unit's line level", vbCritical, "Danger Will Robinson"
        Exit Sub
    End If
    On Error GoTo FixMe
    Dim ExtraFields As String
    If Me.UnitStatusCbo = "A" Then
        ExtraFields = ", Units.UnitDocAccpt = -1, Units.unitstatus = 'A' "
        Call UpdateMassUnits("Units.UnitStatus", Me.UnitStatusCbo, "WHERE (IsDate(Units.UnitCOABack) OR IsDate(Units.UnitCofACopy)) AND Units.UnitStatus <> 'A'", ExtraFields)
    Else
        Call UpdateMassUnits("Units.UnitStatus", Me.UnitStatusCbo)
    End If
    ModMade = -1
    Me.Form.Refresh
    Call CheckMessage
FixMe:
End Sub

Private Sub COASubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitCOASent", Me.COASent)
    ModMade = -1
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub
Private Sub COACopySubmit_Click()
    On Error GoTo FixMe
    Dim ExtraFields As String
    If Me.Parent.LeaseList.Column(2) = "FA" Then
        If IsDate(Me.COACopyCbo) Then
           ExtraFields = ", Units.UnitDocAccpt = 1, Units.unitstatus = 'A' "
           Call UpdateMassUnits("Units.UnitCofACopy", Me.COACopyCbo, "WHERE Units.UnitStatus ='B' or 'A'", ExtraFields)
        Else
           ExtraFields = ", Units.UnitDocAccpt = 0, Units.unitstatus = 'B' "
           Call UpdateMassUnits("Units.UnitCofACopy", Me.COACopyCbo, "WHERE Units.UnitStatus ='A' AND Not IsDate(Units.UnitCofACopy ", ExtraFields)
        End If
    End If
    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub
Private Sub POASubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitPOA", Me.POASent)
    ModMade = -1
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub
Private Sub TitleAppSubmit_Click()
    'On Error GoTo FixMe
    Dim RST As DAO.Recordset
    Dim Tx, x As Integer
    Dim Committ As Variant
    SQLx = "SELECT W.UnitID, U.UnitTASent, T.UnitTitleLogID, T.UnitTitleLogDate, T.EntryReasonId " & _
        "FROM (Bulk_Update_WorkTable as W " & _
        "INNER JOIN Units As U ON W.UnitID = U.UnitID) " & _
        "LEFT JOIN UnitTitleLog AS T ON W.UnitID = T.UnitRef " & _
        "WHERE (((T.EntryReasonId)=18));"
    Set RST = CurrentDb.OpenRecordset(SQLx, dbOpenSnapshot)
    On Error Resume Next
    RST.MoveLast
    RST.MoveFirst
    Tx = RST.RecordCount
    On Error GoTo 0
    If Tx >= 1 Then
        Committ = MsgBox("You're about to CHANGE the App Set Date for ALL units listed." & vbNewLine & _
                         "This includes the Title Log entry for the Title Application." & vbNewLine & _
                         "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Are you serious!, Change THE Date?")
    Else
        Committ = MsgBox("This will only update units that do NOT have a a Title Application Date set." & vbNewLine & _
            "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Set Date?")
    End If
    If Committ = vbNo Then
        RST.Clone
        Set RST = Nothing
        Exit Sub
    End If
    ModMade = -1
    If Tx >= 1 Then
        For x = 1 To Tx
            DoCmd.SetWarnings False
            SQLz = "Update Units SET UnitTASent = '" & Me.TitleAppSent & "' " & _
                   "WHERE UNITID = " & RST!UnitID
            DoCmd.RunSQL (SQLz)
            If RST!UnitTitleLogId <> Null Then
                SQLz = "Update UnitTitleLog SET UnitTitleLogDate = '" & Me.TitleAppSent & "' " & _
                       "WHERE UnitTitleLogID = " & RST!UnitTitleLogId
                DoCmd.RunSQL (SQLz)
            End If
            DoCmd.SetWarnings True
            RST.MoveNext
        Next x
    Else
        SQLz = "UPDATE Units INNER JOIN Bulk_Update_WorkTable ON Units.UnitID = Bulk_Update_WorkTable.UnitID " & _
            "SET Units.UnitTASent = '" & Me.TitleAppSent & "' " & _
            "WHERE Not UnitTASent Is Null;"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)

        SQLz = "INSERT INTO UnitTitleLog ( UnitRef, EntryTypeId, EntryReasonId, EntryReason, UnitTitleLogDate, ExpectedReturn, UserID ) " & _
            "SELECT Bulk_Update_WorkTable.UnitID, 2, 18, 'Title Application Sent', #" & Me.TitleAppSent & "#, -1, '" & FAUserName & "' " & _
            "FROM Bulk_Update_WorkTable LEFT JOIN UnitTitleAppSentQ ON Bulk_Update_WorkTable.UnitID = UnitTitleAppSentQ.UnitRef " & _
            "WHERE (((UnitTitleAppSentQ.UnitRef) Is Null));"

        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
    End If
FixMe:
    RST.Clone
    Set RST = Nothing
    Me.Form.Refresh
    DoCmd.SetWarnings True
    Exit Sub
End Sub
Private Sub CheckMessage()
If Me.Parent.LeaseList.Column(2) = "FA" And AllUnitsInGroup_Accepted(Me.GroupRef) = True Then Call AllUnitsInGroup_AcceptedMessage(Me.GroupRef)
End Sub
Private Sub UnitTASent_AfterUpdate()
    If IsDate(Me.UnitTASent) Then
        If DLookup("EntryReasonId", "UnitTitleLog", "UnitRef=" & Me.UnitID) = 18 Then
            MsgBox "A Title Application Date has already be set for Unit " & Me.UnitID & "." & vbNewLine & vbNewLine & "If you need to send for a new or corrected title, " & vbNewLine & "Please enter a new title log below." & vbNewLine & vbNewLine & "Date will revert (not saved)", vbCritical, "Title App Already set"
            Me.Undo
            Exit Sub
        End If
        SQLz = ""
        SQLz = SQLz & "INSERT INTO UnitTitleLog ( UnitRef,  EntryReason,  UnitTitleLogDate,  EntryTypeId,  EntryReasonId,  ExpectedReturn,  UserID )"
        SQLz = SQLz & "SELECT " & Me.UnitID & ", 'Title Application Sent', #" & Me.UnitTASent & "#, 2, 18, True, '" & FAUserName & "';"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        ModMade = -1
        DoCmd.SetWarnings True
    End If
End Sub
Private Sub UnitID_DblClick(Cancel As Integer)
    Call OpenUnitsForm(Me.UnitID)
    Me.UnitUnitNum.SetFocus
End Sub
