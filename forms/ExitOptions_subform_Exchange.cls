Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public SQLz, SCHGrpX, Response, SQLset As String
Public Sub ClearBoxes()
    Dim Ccrl As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Medit" Then
            Me(Ccrl.Name).Value = ""
        End If
    Next Ccrl
End Sub


Private Sub DeleteMe_Click()
    Response = MsgBox("You're about to Delete Unit# " & Me.unitunitnum & "'s " & vbNewLine & Me.OptionTypeDesc & " option entry , " & vbNewLine & "Are you sure?", vbQuestion + vbYesNo + vbCritical, "What? Delete!")

    SQLset = " WHERE UnitID = " & Me.UnitID & "and OptionTypeID = 1;"

    If Response = vbNo Then Exit Sub
    On Error GoTo FixMe
    SQLz = "Delete From EXITOptions_WorkTable WHERE UnitID = " & Me.UnitID
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    SQLz = "Delete From UnitEXITOption WHERE LeaseOptionID = " & Me.LeaseOptionID & ";"
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Me.Form.Requery
    Me.Parent.Units_List.Requery
FixMe:
    Me.Form.Refresh
End Sub

Private Sub EndOrActualDate_Change()
    Call SetModified
End Sub



Private Sub FinalRentDate_Change()
Call SetModified
End Sub

Private Sub Form_Current()
    Me.ReplacedByGroupCBO.Requery
    Me.ReplacedByGroupID.Requery
    Me.ReplacedByUnitID.Requery
    Me.txtCurrent = Me.UnitID
End Sub

'Private Sub Option_BeforeUpdate(Cancel As Integer)
'    If IsNull(Me.Option) Then
'        SQLz = "INSERT INTO UnitExitOption ( UnitId, GroupId, OptionSelected, OptionTerm, ExchangePool, DateSeleted ) " & _
'               "SELECT " & Me.UnitID & ", " & Me.GroupRef & "," & Me.Option.Text & ",0, 0, #" & Now() & "# ;"
'        DoCmd.RunSQL (SQLz)
'    End If
'End Sub

Public Sub ConfrimIT(ByVal FieldBoX As String, ByVal TragetX As String)
    Dim Committ As String
    If Me(FieldBoX).Value & "TEST" = "TEST" Then
        Committ = MsgBox("You're about set the " & TragetX & " to Null, for ALL units listed." & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Clear " & TragetX & "?")
        If Committ = vbNo Then Exit Sub
    Else
        Committ = MsgBox("You're about set the " & TragetX & " to " & Me(FieldBoX).Value & ", for ALL units listed." & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Change " & TragetX & "?")
    End If
End Sub

Private Sub ModifyEXTD_Click()
DoCmd.RunCommand acCmdSaveRecord
    Form_EXITOptions.RemoteSelectSingleUnit (Me.UnitID)
    Form_EXITOptions.ExtendSelected_Click
End Sub

Private Sub ModifyTERM_Click()
DoCmd.RunCommand acCmdSaveRecord
    Form_EXITOptions.RemoteSelectSingleUnit (Me.UnitID)
    Form_EXITOptions.TerminateSelected_Click
End Sub

Private Sub Notes_Change()
    Call SetModified
End Sub

Private Sub NotesSubmit_Click()
    If Me.NotesCbo & "test" = "test" Then
        Exit Sub
    End If
    SQLset = "SET Notes = '" & Me.NotesCbo & "', UserName='" & FaUserName & "', dateposted='" & Now() & "'WHERE OptionTypeID = 1;"
    Response = MsgBox("You're about set the Notes for all selected units, Are you sure?", vbQuestion + vbYesNo + vbCritical, "Set Notes?")
    If Response = vbNo Then Exit Sub
    On Error GoTo FixMe
    SQLz = "UPDATE UnitExitOption INNER JOIN EXITOptions_WorkTable ON UnitExitOption.UnitID = EXITOptions_WorkTable.UnitID  " & SQLset
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
FixMe:
    DoCmd.SetWarnings True
    Me.Form.Refresh
    Exit Sub
End Sub


Private Sub ReplacedByGroupID_Change()
Call SetModified
End Sub

Private Sub ReplacedByGroupID_Click()
If SCHGrpX <> Me.ReplacedByGroupID.Column(1) And Not IsNull(Me.ReplacedByUnitID) Then
MsgBox "AssetID: " & Me.UnitID & " has been Exchanged with AssetID: " & Me.ReplacedByUnitID & " " & vbNewLine & _
       "You cannot change to a different Schedule Group" & vbNewLine & vbNewLine & _
       "If you want to change the Schedule Group, follow these steps:" & vbNewLine & "1: Clear the matched AssetID: " & Me.ReplacedByUnitID & vbNewLine & _
       "2: Clear the Estimated and Actual Exchanged dates" & vbNewLine & _
       "3: Select a Schedule Group", vbCritical, "Units Already Exchanged-Matched!"
Me.Undo
Exit Sub
Else
    Me.ReplacedByScheduleID = Me.ReplacedByGroupID.Column(3)
    Me.ReplacedByUnitID.Requery
    Me.StartOrPurposedDate = DLookup("Min(UnitEstDelDate)", "Units", "GroupRef=" & ReplacedByGroupID & "")
    Call SetModified
End If
End Sub

Public Sub ReplacedByGroupID_GotFocus()
If Not IsNull(Me.ReplacedByGroupID.Column(1)) Then SCHGrpX = Me.ReplacedByGroupID.Column(1)
End Sub

Private Sub ReplacedByUnitID_Change()
    If Me.ReplacedByUnitID.Text & "test" = "test" Then
         Me.ReplacedByUnitID = Null
         Me.EndOrActualDate = Null
         Me.StartOrPurposedDate = Null
         Me.UnitPDStartDate = Null
         Me.FinalRentDate = Null
    Else
         Dim BLCDx As Date
         Dim GroupIDi As Long
         Dim AdvArr As String
         Dim RST As Recordset
         DoCmd.RunCommand acCmdSaveRecord
         Set RST = CurrentDb.OpenRecordset("Select * " & _
                   " From vw_EXchangeITDates Where ReplacedByUnitId=" & Me.ReplacedByUnitID & ";")
         BLCDx = RST!oBLCD
         GroupIDi = RST!oGroupID
         AdvArr = DLookup("LeasePaidAdvArr", "SchGrp", "sGrpID=" & GroupIDi)
         If IsDate(RST!oActualOLDate) Or IsDate(Me.StartOrPurposedDate) Then Me.StartOrPurposedDate = Me.StartOrPurposedDate Else Me.StartOrPurposedDate = RST!exStartDate
         If IsDate(RST!oActualOLDate) Then
            If AdvArr = "Arr" Then Me.FinalRentDate = DateAdd("d", 1, DateAdd("m", -0, RST!oActualOLDate))
            If AdvArr = "Adv" Then Me.FinalRentDate = DateAdd("d", 0, DateAdd("m", -1, RST!oActualOLDate))
        Else
            Me.FinalRentDate = RST!exFinalRent
        End If
         Me.EndOrActualDate = Nz(RST!oActualOLDate, RST!exActual)


         Me.UnitPDStartDate = RST!nPDStart
         RST.close
         Set RST = Nothing
    End If
    Call SetModified
    DoCmd.RunCommand acCmdSaveRecord
    ReplacedByUnitID.Requery
End Sub

Private Sub ReplacedByUnitID_DblClick(Cancel As Integer)
    If IsNull(ReplacedByUnitID) Then Exit Sub
    Call OpenUnitsForm(Me.ReplacedByUnitID)
    'DoCmd.OpenForm "UnitsForm", , , "UNITID=" & Me.ReplacedByUnitId & ""
End Sub

Private Sub FinalRentDateSubmit_Click()
    If Me.FinalRentDateCBO & "test" = "test" Then
        Exit Sub
    End If
    SQLset = "SET FinalRentDate = #" & Me.FinalRentDateCBO & "#, " & _
             "UserName='" & FaUserName & "', dateposted='" & Now() & "'WHERE OptionTypeID = 1;"
    Response = MsgBox("You're about set the Final Rent Date for all selected units, Are you sure?", vbQuestion + vbYesNo + vbCritical, "Final Rent Date")
    If Response = vbNo Then Exit Sub
'    On Error GoTo FixMe
    SQLz = "UPDATE UnitExitOption INNER JOIN EXITOptions_WorkTable ON UnitExitOption.UnitID = EXITOptions_WorkTable.UnitID  " & SQLset
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
FixMe:
    DoCmd.SetWarnings True
    Me.Form.Refresh
End Sub
Private Sub ReplcedBySubmit_Click()
    If Me.ReplacedByGroupCBO & "test" = "test" Then
        Exit Sub
    End If
    SQLset = "SET ReplacedByScheduleID = " & Me.ReplacedByGroupCBO.Column(3) & ", ReplacedByGroupID = " & Me.ReplacedByGroupCBO.Column(0) & ", ReplacedByUnitID=Null, " & _
             "UserName='" & FaUserName & "', dateposted='" & Now() & "'WHERE OptionTypeID = 1;"
    Response = MsgBox("You're about set the Replacement Schedule and Group for all selected units, Are you sure?", vbQuestion + vbYesNo + vbCritical, "Clear Replacements?")
    If Response = vbNo Then Exit Sub
    On Error GoTo FixMe
    SQLz = "UPDATE UnitExitOption INNER JOIN EXITOptions_WorkTable ON UnitExitOption.UnitID = EXITOptions_WorkTable.UnitID  " & SQLset
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
FixMe:
    DoCmd.SetWarnings True
    Me.Form.Refresh
End Sub
'Private Sub PurposedSubmit_Click()
'    If Me.StartOrPurposedDateCBO & "test" = "test" Then
'        Exit Sub
'    End If
'
'    SQLset = "SET StartOrPurposedDate = '" & Me.StartOrPurposedDateCBO & "', UserName='" & FAUserName & "', dateposted='" & Now() & "'WHERE OptionTypeID = 1;"
'    Response = MsgBox("You're about set the Purposed ExchangeIT date for all selected units, Are you sure?", vbQuestion + vbYesNo + vbCritical, "Set Date?")
'    If Response = vbNo Then Exit Sub
'    On Error GoTo Fixme
'    SQLz = "UPDATE UnitExitOption INNER JOIN EXITOptions_WorkTable ON UnitExitOption.UnitID = EXITOptions_WorkTable.UnitID  " & SQLset
'    DoCmd.SetWarnings False
'    DoCmd.RunSQL (SQLz)
'Fixme:
'    DoCmd.SetWarnings True
'    Me.Form.Refresh
'    Exit Sub
'End Sub
Private Sub ActualDateSubmit_Click()
    If Me.EndOrActualDateCBO & "test" = "test" Then Exit Sub

    SQLset = "SET EndOrActualDate = '" & Me.EndOrActualDateCBO & "', UserName='" & FaUserName & "', dateposted='" & Now() & "'WHERE OptionTypeID = 1;"
    Response = MsgBox("You're about set the ACTUAL ExchangeIT date for all selected units, Are you sure?", vbQuestion + vbYesNo + vbCritical, "Set Date?")
    If Response = vbNo Then Exit Sub
    On Error GoTo FixMe
    SQLz = "UPDATE UnitExitOption INNER JOIN EXITOptions_WorkTable ON UnitExitOption.UnitID = EXITOptions_WorkTable.UnitID  " & SQLset
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
FixMe:
    DoCmd.SetWarnings True
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub SaveUpdateMe_Click()
Me.Form.Refresh
Call MakeExitOptionLocal(False, True)
End Sub

Private Sub Send2Excel_Click()
    Dim XLS As String
    XLS = "|D,HH|I,HH|F,BC|H,$E|J,BC|M,$E|Q,$E|"
    SendFormData2Excel Me, "Units", "Buy Out Info for " & Me.UnitCount & " Units", True, XLS, Me.UnitCount
End Sub

Private Sub StartOrPurposedDate_AfterUpdate()
    Call SetModified
End Sub

Private Sub StartOrPurposedDate_Change()
Call SetModified
End Sub

Private Sub UnitID_DblClick(Cancel As Integer)
    Call OpenUnitsForm(Me.UnitID)
    'DoCmd.OpenForm "UnitsForm", , , "UNITID=" & Me.UnitID & ""
    Forms!Unitsform.UnitsTabControl.Value = 8
    'Me.UnitNum.SetFocus
End Sub
Public Sub SetModified()
    Me.UserName = FaUserName
    Me.DatePosted = Now
End Sub

Private Sub UnitPDStartDate_Change()
Call SetModified
End Sub
