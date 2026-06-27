Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz, SQLx As String
Dim PrevLocID As Long

Private Sub Form_AfterUpdate()
   Me.Parent.ModMade = -1
End Sub

Private Sub Form_Current()
On Error GoTo whoops
   Me.txtCurrent = Me.UnitID
whoops:
   Exit Sub
End Sub
Private Sub GoHome_Click()
DoCmd.OpenForm " FAMainMenu"
End Sub
Private Sub Form_ApplyFilter(Cancel As Integer, ApplyType As Integer)
    Dim zFilter As String
    On Error GoTo FixMe    ' if the user is just sorting
    If Me.Form.Filter Like "*LocationName*" Then
        Dim sLookUp As String
        Dim iLocID As Long
        sLookUp = SplitString(SplitString(Me.Form.Filter, "=", 1), ")", 0)
        sLookUp = Mid(sLookUp, 2, Len(sLookUp) - 2)
        iLocID = DLookup("LocationID", "Locations", "ClientGroupId=" & Me.Parent.ClientGroupID & " AND LocationName='" & sLookUp & "'")
        zFilter = Replace(SplitString(SplitString(Me.Form.Filter, ".", 0), "[LookUp_", 1), "]", "") & "=" & iLocID
    Else
        zFilter = Replace(Replace(Replace(Me.Form.Filter, "[Bulk_Update_SelectedUnits].", ""), "]", ""), "[", "")
    End If
    If zFilter <> "" Then Call Me.Parent.SetUnits(zFilter)
    Call Me.Parent.SetOrderBy(Me.OrderBy)
    Me.Form.Filter = ""
FixMe:
End Sub
Public Sub ClearBoxes()
Dim Ccrl As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Medit" Then
            Me(Ccrl.Name).Value = ""
        End If
     Next Ccrl
End Sub
Private Sub DeliverySubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitDeliveryLocationRef", Me.DeliveryCbo)
    ModMade = -1
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub
Public Sub ConfrimIT(ByVal FieldBoX As String, ByVal TragetX As String)
    Dim Committ As String
    If Me(FieldBoX).Value & "TEST" = "TEST" Then
        Committ = MsgBox("You're about set the " & TragetX & " to Null, for ALL units listed." & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Clear " & TragetX & "?")
        If Committ = vbNo Then Exit Sub
    Else
        Committ = MsgBox("You're about set the " & TragetX & " to " & Me(FieldBoX).Value & ", for ALL units listed." & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Change " & TragetX & "?")
    End If
End Sub

Private Sub Form_Dirty(Cancel As Integer)
   Me.Parent.ModMade = -1
End Sub



Private Sub UnitGarageLocationRef_DblClick(Cancel As Integer)
  DoCmd.close acForm, "LocationAddressesEdit"
  DoCmd.OpenForm "LocationAddressesEdit", acNormal, , , , , DLookup("ClientGroupID", "Locations", "LocationID=" & UnitGarageLocationRef & "") & ":" & UnitGarageLocationRef
End Sub

Private Sub UnitGarageLocationRef_GotFocus()
PrevLocID = Nz(Me.UnitGarageLocationRef.Column(0), 0)
End Sub


Private Sub UnitGarageLocationRef_AfterUpdate()
If Not Nz(DLookup("Locationsn", "Locations", "LocationID=" & PrevLocID), "") Like "*TBD*" Then Call GarageChange(Me.UnitID, True)
End Sub

Private Sub GarageSubmit_Click()
        Dim rc As Integer
        SQLx = ""
        SQLx = "SELECT AssetID, LocationSn FROM vw_Sixkeys "
        SQLx = SQLx & " WHERE AssetID IN (Select UnitID From Bulk_Update_WorkTable)"
        SQLx = SQLx & " AND LocationSn <> 'TBD'"
        Set RST = CurrentDb.OpenRecordset(SQLx, dbOpenSnapshot)
        SQLx = ""
        On Error Resume Next
        RST.MoveLast
        RST.MoveFirst
        On Error GoTo 0
        rc = RST.RecordCount
        RST.close
        Set RST = Nothing
        If rc > 0 Then Call GarageChange(, True) Else Call GarageChange(, False)
End Sub
Public Function GarageChange(Optional UnitIDi As Long, Optional AskReason As Boolean)
  '  On Error GoTo Fixme
    Dim i, isBUnit As Integer
    If DMin("UnitStatus", "Units", "GroupREF=" & Me.Parent.GrpID) = "B" Then isBUnit = 1 Else isBUnit = 0
    If isBUnit = 0 And Me.Parent.LeaseList.Column(2) = "FA" Then
        Dim RST As DAO.Recordset
        Dim MoveDate, MoveReason, NewLoc As String
        Dim msg1 As String
        SQLx = ""
        SQLx = "SELECT AssetID, LocationID FROM vw_Sixkeys "
        If UnitIDi <> 0 Then
            NewLoc = Me.UnitGarageLocationRef.Column(0)
            SQLx = SQLx & " WHERE AssetId=" & UnitIDi
        Else
            NewLoc = Me.GarageCbo
            SQLx = SQLx & " WHERE AssetID IN (Select UnitID From Bulk_Update_WorkTable)"
        End If
        If AskReason = True Then
        For i = 1 To 2
AskReason:
  If isBUnit = 1 Then GoTo UpdateGarage
            msg1 = ""
            If UnitIDi <> 0 Then
                msg1 = "<B>Moving FA Leased Units!</B><br/><br/>"
                msg1 = msg1 & "You are about to Move<b> Asset: " & Me.UnitID & "</b> to:<br/>"
                msg1 = msg1 & "<b>" & Me.UnitGarageLocationRef.Column(1) & "</b><br/><br/>"
            Else
                msg1 = "<B>Moving FA Leased Units!</B><br/><br/>"
                msg1 = msg1 & "You are about to Move Units to:<br/>"
                msg1 = msg1 & "<b>" & Me.GarageCbo.Column(1) & "</b><br/><br/>"
            End If
            If i = 1 Then
                msg1 = msg1 & "Provide the<b> Move DATE</b><i> 'mm/dd/yyyy'</i> below:<br/>"
                msg1 = msg1 & "<i>you may type 'STOP' to cancel this move</i>"
                MoveDate = OpenInput(msg1, "Question", AtlaasVersion, "Move Date", "Move Date")
                If MoveDate = "stop" Then GoTo UndoMe
                If Not IsDate(Trim(MoveDate)) = True Then GoTo AskReason
            End If
            If i = 2 Then
                msg1 = msg1 & "Provide the<b> Reason</b> below:<br/>"
                msg1 = msg1 & "<i>you may type 'STOP' to cancel this move</i>"
                MoveReason = OpenInput(msg1, "Question", AtlaasVersion, "Move Reason", "Move Reason")
                If MoveReason & "test" = "test" Then GoTo AskReason
                If MoveReason = "stop" Or MoveReason = "cancel" Then GoTo UndoMe
            End If
        Next i
        i = 0
        Set RST = CurrentDb.OpenRecordset(SQLx, dbOpenSnapshot)
        On Error Resume Next
        RST.MoveLast
        RST.MoveFirst
        On Error GoTo 0
        'Call MessageUser("Update", "Getting AssetIDs", "working...")
        For i = 1 To RST.RecordCount
            SQLz = "Insert Into LocationChangeHistory (AssetId, OldLocationId, NewLocationId, LocationChangeDate, GroupId, Reason,FAUserName, LocationType) " & _
                   "Values (" & RST!AssetID & ", " & RST!LocationID & ", " & NewLoc & ", '" & MoveDate & "'," & Me.Parent.GroupList & ",'" & MoveReason & "','" & FAUserName & "','Garage')"
            DoCmd.SetWarnings False
            DoCmd.RunSQL (SQLz)
            DoCmd.SetWarnings True
            RST.MoveNext
        Next i
        RST.close
        Set RST = Nothing
        End If
    End If
UpdateGarage:
    If isBUnit = 1 Then Exit Function
    If UnitIDi <> 0 Then GoTo FixMe
    SQLz = "UPDATE Units INNER JOIN Bulk_Update_WorkTable ON Units.UnitID = Bulk_Update_WorkTable.UnitID SET Units.UnitGarageLocationRef = " & Me.GarageCbo & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    ModMade = -1
    DoCmd.SetWarnings True
FixMe:
    Me.Form.Refresh
    Exit Function
UndoMe:
    If UnitIDi <> 0 Then Me.UnitGarageLocationRef = PrevLocID
End Function

Private Sub TitleSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitTitleLocationRef", Me.TitleCbo)
    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub
Private Sub RegistrationSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitRegLocationRef", Me.RegistrationCbo)
    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub
Private Sub IRPSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitIRP", Me.IRPCBO)
    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub

Private Sub RegExpireSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitRegExp", Me.RegExpire)
    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub

Private Sub ExemptSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitTaxExempt", Me.ExemptCbo)
    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub

Private Sub IsUpFrontSubmit_Click()
    On Error GoTo FixMe
If Me.IsUpFrontCbo & "test" = "test" Then
    Call UpdateMassUnits("Units.IsUpFront", Null)
Else
    Call UpdateMassUnits("Units.IsUpFront", Me.IsUpFrontCbo)
End If
    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub

Private Sub UnitID_DblClick(Cancel As Integer)
Call OpenUnitsForm(Me.UnitID)
Me.UnitUnitNum.SetFocus
End Sub

Private Sub WeightSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitRegW", Me.Weight)
    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub
Private Sub NoteSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitTaxNotes", Me.Notes)
    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub
