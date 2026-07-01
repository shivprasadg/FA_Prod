Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Dim SQLz, MissingX, Committ As String
Public RST    As DAO.Recordset

Private Sub EntryReasonId_Click()
   Me.EntryReason = Me.EntryReasonId.Column(1)
   Me.EntryTypeID = Me.EntryReasonId.Column(2)
   Me.ExpectedReturn = Me.EntryReasonId.Column(3)
End Sub

Private Sub EntryReasonSubmit_Click()
    If Me.EntryReasonCbo.Column(1) & "test" = "test" Then Exit Sub
    Call ConfrimIT("Me.EntryReasonCbo.Column(1)", "Title Log Reason")
    SQLz = "UPDATE RMKTUnitTitleLog_Work "
    SQLz = SQLz & "SET EntryReasonID =" & Me.EntryReasonCbo.Column(0)
    SQLz = SQLz & ", EntryReason ='" & Me.EntryReasonCbo.Column(1) & "'"
    SQLz = SQLz & ", EntryTypeId =" & Me.EntryReasonCbo.Column(2)
    SQLz = SQLz & ", ExpectedReturn =" & Me.EntryReasonCbo.Column(3)
    On Error GoTo FixMe
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
FixMe:
    Me.Form.Refresh
End Sub

Private Sub Form_Current()
On Error Resume Next
    Me.txtCurrent = Me.AssetID
    'Public RST    As DAO.Recordset
    SQLz = "Select ClientShortname As Client, Schedule, UnitGroup From vw_SixKeys WHERE AssetID=" & Me.UnitID & ";"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    Me.Parent.BuyNo = "Client: " & RST!Client & " Sch|Grp: " & RST!Schedule & "|" & RST!UnitGroup
    RST.close
    Set RST = Nothing
On Error GoTo 0
End Sub
Public Sub ClearBoxes()
Dim Ccrl As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Medit" Then
            Me(Ccrl.Name).Value = ""
        End If
     Next Ccrl
End Sub

Private Sub LogALLEvents_Click()
    Committ = MsgBox("You're about Insert a 'Title Event' for ALL of untis listed." & _
    "" & vbNewLine & "  > If the Entry is NOT correct it will NOT save and remain." & _
    "" & vbNewLine & "  > If it is valid, it will SAVE and be removed from this list." & _
    "" & vbNewLine & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Set " & TragetX & "?")
    If Committ = vbNo Then Exit Sub

    SQLz = "INSERT INTO UnitTitleLog ( UnitRef, EntryTypeId, EntryReasonId, EntryReason, TitleNumber, TitleSt, UnitTitleLogDate, ExpectedReturn, UserID )"
    SQLz = SQLz & "SELECT AssetID, EntryTypeId, EntryReasonId"
    SQLz = SQLz & ", EntryReason, TitleNumber, TitleSt"
    SQLz = SQLz & ", UnitTitleLogDate, ExpectedReturn, UserID "
    SQLz = SQLz & "From RMKTUnitTitleLog_Work "
    SQLz = SQLz & "WHERE (((EntryReasonId) Is Not Null) "
    SQLz = SQLz & "AND ((nz(LastEntryTypeId,0))<>EntryTypeId));"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True

    SQLz = "Delete From RMKTUnitTitleLog_Work "
    SQLz = SQLz & "Where AssetId in ("
    SQLz = SQLz & "Select AssetID From RMKTUnitTitleLog_Work "
    SQLz = SQLz & "WHERE (((EntryReasonId) Is Not Null) "
    SQLz = SQLz & "And ((LastEntryTypeId) <> EntryTypeId)))"

    Me.Form.RecordSource = ""

    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    Me.Form.RecordSource = "RMKTUnitTitleLog_Work"
    DoCmd.SetWarnings True
End Sub

Private Sub LogDateSubmit_Click()
    If Me.LogDateCbo & "test" = "test" Then Exit Sub
    Call ConfrimIT("LogDateCbo", "Title Log Date")
    SQLz = "UPDATE RMKTUnitTitleLog_Work SET UnitTitleLogDate =#" & Me.LogDateCbo & "#;"
    On Error GoTo FixMe
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
FixMe:
    Me.Form.Refresh
End Sub
Private Sub TitleStSubmit_Click()
    If Me.TitleStCbo & "test" = "test" Then Exit Sub
    Call ConfrimIT("TitleStCbo", "Title State")
    SQLz = "UPDATE RMKTUnitTitleLog_Work SET TitledState ='" & Me.TitleStCbo & "';"
    On Error GoTo FixMe
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
FixMe:
    Me.Form.Refresh
End Sub
Public Sub ConfrimIT(ByVal FieldBoX As String, ByVal TragetX As String)
    Dim Fieldx As String
    If Nz(FieldBoX, "") Like "*column*" Then Fieldx = Me.EntryReasonCbo.Column(1) Else Fieldx = Me(FieldBoX).Value
           Committ = MsgBox("You're about set the " & TragetX & " to" & vbNewLine & "'" & Fieldx & "'" & vbNewLine & "for ALL units listed." & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Set " & TragetX & "?")
    If Committ = vbNo Then Exit Sub
End Sub

Private Sub LogTitleEvent_Click()
Dim WrongDirection, DirX  As String
If Me.EntryTypeID = Null Then Exit Sub
If Me.EntryTypeID = 1 And Me.LastDirection = "In" Then WrongDirection = "You're Trying to Log IN a Title That is Already 'IN'"
If Me.EntryTypeID = 2 And Me.LastDirection = "Out" Then WrongDirection = "You're Trying to Log OUT a Title That is Already 'OUT'"

    If WrongDirection <> "" Then
        MsgBox WrongDirection & vbNewLine & "No action was taken", vbCritical, "You got peanut butter in my chocolate"
        Exit Sub
    End If
    Call CheckFields
    If MissingX <> "You're Missing:" Then
        MsgBox MissingX, vbCritical, "Data missing"
        Exit Sub
    End If
    SQLz = "INSERT INTO UnitTitleLog ( Unitref, TitleNumber, TitleSt, EntryReason, UnitTitleLogDate, EntryTypeId, UserID, EntryReasonID) "
    SQLz = SQLz & "SELECT " & Me.UnitID.Value & " as unitid, '" & Me.UnitTitleNumber.Value & "' AS Tnum"
    SQLz = SQLz & ", '" & Me.TitledState & "' AS State, '" & Me.EntryReasonId.Column(1) & "' AS reason, '" & Me.UnitTitleLogDate & "' AS UnitTitleLogDate "
    SQLz = SQLz & ", 1,'" & FaUserName & "', " & Me.EntryReasonId.Column(0) & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    'Call UpdateUnitTitleState
    'Call UpdateUnitTitleNum
    If Me.EntryTypeID = 1 Then DirX = "In" Else DirX = "Out"
    MsgBox "Title # " & Me.UnitTitleNumber & " has been checked " & DirX & vbNewLine & " The unit Will be removed from the list", , " Checked In"
    DoCmd.SetWarnings False
    SQLz = "Delete from RMKTUnitTitleLog_Work Where AssetID=" & Me.UnitID.Value
    Me.LogDateCbo.SetFocus
    Me.Form.RecordSource = ""
    DoCmd.RunSQL (SQLz)
    Me.Form.RecordSource = "RMKTUnitTitleLog_Work"
    DoCmd.SetWarnings True
End Sub

Public Sub CheckFields()
    MissingX = "You're Missing:"
    If Me.EntryReasonId & "test" = "test" Then MissingX = MissingX & vbNewLine & " >  Reason"
    If Me.UnitTitleNumber & "test" = "test" Then MissingX = MissingX & vbNewLine & " >  Title Number"
    If Me.TitledState & "test" = "test" Then MissingX = MissingX & vbNewLine & " >  Titled State"
    If Me.UnitTitleLogDate & "test" = "test" Then MissingX = MissingX & vbNewLine & " >  Title Log Date"
    If MissingX <> "You're Missing:" Then MissingX = MissingX & vbNewLine & "Exiting... No actions were logged"
End Sub



Private Sub UnitID_DblClick(Cancel As Integer)
    Call OpenUnitsForm(Me.AssetID)
End Sub
