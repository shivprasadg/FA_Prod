Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public SQLz As String
Public Sub ClearBoxes()
    Dim Ccrl As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Medit" Then
            Me(Ccrl.Name).Value = ""
        End If
    Next Ccrl
End Sub


Private Sub AssigneeNotified_AfterUpdate()
    Call SetModified
End Sub

Private Sub AssigneeReplied_AfterUpdate()
    Call SetModified
End Sub

Private Sub DeleteMe_Click()
    Dim Response As String
    Response = MsgBox("You're about to Delete Unit# " & Me.UnitUnitNum & "'s " & vbNewLine & Me.OptionSubType & " option entry , " & vbNewLine & "Are you sure?", vbQuestion + vbYesNo + vbCritical, "What? Delete!")

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
    Exit Sub
End Sub

Private Sub EndOrActualDate_AfterUpdate()
    Call SetModified
End Sub

Private Sub Fees_AfterUpdate()
    Call SetModified
End Sub

Private Sub FinalRentDateSubmit_Click()
    If Me.FinalRentDateCBO & "test" = "test" Then
        Exit Sub
    End If
    Dim Response, SQLset As String
    SQLset = "SET FinalRentDate = '" & Me.FinalRentDateCBO & "', UserName='" & FAUserName & "', dateposted='" & Now() & "'WHERE OptionTypeID = 2;"
    Response = MsgBox("You're about set the Final Rent Date for all selected units, Are you sure?", vbQuestion + vbYesNo + vbCritical, "Set Final Rent Date?")
    If Response = vbNo Then Exit Sub
    On Error GoTo FixMe
    SQLz = "UPDATE UnitExitOption INNER JOIN EXITOptions_WorkTable ON UnitExitOption.UnitID = EXITOptions_WorkTable.UnitID  " & SQLset
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub FinalPaymntPaid_AfterUpdate()
    Call SetModified
End Sub

Private Sub ModifyExch_Click()
    DoCmd.RunCommand acCmdSaveRecord
    Form_EXITOptions.RemoteSelectSingleUnit (Me.UnitID)
    Form_EXITOptions.ExchangeSelected_Click
End Sub

Private Sub ModifyEXTD_Click()
    DoCmd.RunCommand acCmdSaveRecord
    Form_EXITOptions.RemoteSelectSingleUnit (Me.UnitID)
    Form_EXITOptions.ExtendSelected_Click
End Sub

Private Sub Notes_AfterUpdate()
    Call SetModified
End Sub

Private Sub OptionSubTypeSubmit_Click()
    If Me.OptionSubTypeCbo & "test" = "test" Then
        Exit Sub
    End If
    Dim Response, SQLset As String
    SQLset = "SET OptionSubType = '" & Me.OptionSubTypeCbo & "', UserName='" & FAUserName & "', dateposted='" & Now() & "'WHERE OptionTypeID = 2;"
    Response = MsgBox("You're about set the end Of Lease Type for all selected units, Are you sure?", vbQuestion + vbYesNo + vbCritical, "Set Type?")
    If Response = vbNo Then Exit Sub
    On Error GoTo FixMe
    SQLz = "UPDATE UnitExitOption INNER JOIN EXITOptions_WorkTable ON UnitExitOption.UnitID = EXITOptions_WorkTable.UnitID  " & SQLset
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

'Private Sub PaymntPaidSubmit_Click()
'    If Me.FinalPaymntPaidCBO & "test" = "test" Then
'        Exit Sub
'    End If
'    Dim Response, SQLset As String
'    SQLset = "SET FinalPaymntPaid = '" & Me.FinalPaymntPaidCBO & "', UserName='" & FAUserName & "', dateposted='" & Now() & "'WHERE OptionTypeID = 2;"
'    Response = MsgBox("You're about set the Final Payment Date for all selected units, Are you sure?", vbQuestion + vbYesNo + vbCritical, "Set Paid Date?")
'    If Response = vbNo Then Exit Sub
'    On Error GoTo Fixme
'    SQLz = "UPDATE UnitExitOption INNER JOIN EXITOptions_WorkTable ON UnitExitOption.UnitID = EXITOptions_WorkTable.UnitID  " & SQLset
'    DoCmd.SetWarnings False
'    DoCmd.RunSQL (SQLz)
'    DoCmd.SetWarnings True
'Fixme:
'    Me.Form.Refresh
'    Exit Sub
'End Sub
Private Sub Form_Current()
On Error Resume Next
    Me.txtCurrent = Me.UnitID
On Error GoTo 0
    ' Me.ReplacedByUnitId.Requery
    ' Me.ReplacedByGroupCBO.Requery
    ' Me.ReplacedByGroupID.Requery
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

Private Sub Residual_AfterUpdate()
    Call SetModified
End Sub

Private Sub SaveUpdateMe_Click()
Call MakeExitOptionLocal(False, True)
Form_EXITOptions.SetGapLapData (2)
Me.Form.Refresh
End Sub

Private Sub StartOrPurposedDateCBO_AfterUpdate()
'   If Me.TermCBO & "test" <> "test" Then Me.EndOrActualDateCBO = DateAdd("M", Me.TermCBO, Me.StartOrPurposedDateCBO) - 1
'   If Me.TermCBO & "test" = "test" And IsDate(Me.EndOrActualDateCBO) Then Me.TermCBO = DateDiff("M", Me.StartOrPurposedDateCBO, Me.EndOrActualDateCBO)
End Sub

Private Sub TermCBO_Click()
   ' If IsNull(Me.TermCBO) Then Exit Sub
   ' If IsDate(Me.StartOrPurposedDateCBO) Then Me.EndOrActualDateCBO = DateAdd("M", Me.TermCBO, Me.StartOrPurposedDateCBO) - 1
End Sub
Private Sub Term_Click()
    'If Me.Term & "test" = "test" Then
    '    Exit Sub
    'End If
    'If IsDate(Me.StartOrPurposedDate) Then Me.EndOrActualDate = DateAdd("M", Me.Term, Me.StartOrPurposedDate) - 1
End Sub
Private Sub StartOrPurposedDate_AfterUpdate()
' If Me.Term & "test" <> "test" Then Me.EndOrActualDate = DateAdd("M", Me.Term, Me.StartOrPurposedDate) - 1
' If Me.Term & "test" = "test" And IsDate(Me.EndOrActualDate) Then Me.TermCBO = DateDiff("M", Me.StartOrPurposedDate, Me.EndOrActualDate)
    Call SetModified
End Sub

Private Sub TerminateEarly_AfterUpdate()
    Call SetModified
End Sub

'Private Sub TermSubmit_Click()
'    If Me.TerminateEarlyCBO & "test" = "test" Then
'        Exit Sub
'    End If
'    Dim Response, SQLset As String
'    SQLset = "SET TerminateEarly = " & Me.TerminateEarlyCBO & ", "
'    SQLset = SQLset & "UserName='" & FAUserName & "', dateposted='" & Now() & "'WHERE OptionTypeID = 2;"
'    Response = MsgBox("You're about set the Term (Months) for all selected units, Are you sure?", vbQuestion + vbYesNo + vbCritical, "Set Extended Term?")
'    If Response = vbNo Then Exit Sub
'    On Error GoTo Fixme
'    SQLz = "UPDATE UnitExitOption INNER JOIN EXITOptions_WorkTable ON UnitExitOption.UnitID = EXITOptions_WorkTable.UnitID  " & SQLset
'    DoCmd.SetWarnings False
'    DoCmd.RunSQL (SQLz)
'    DoCmd.SetWarnings True
'Fixme:
'    Me.Form.Refresh
'    Exit Sub
'
'End Sub
Private Sub PurposedSubmit_Click()
    If Me.StartOrPurposedDateCBO & "test" = "test" Then
        Exit Sub
    End If
    Dim Response, SQLset As String
    SQLset = "SET StartOrPurposedDate = '" & Me.StartOrPurposedDateCBO & "', UserName='" & FAUserName & "', dateposted='" & Now() & "'WHERE OptionTypeID = 2;"
    Response = MsgBox("You're about set the Projected Termination Date for all selected units, Are you sure?", vbQuestion + vbYesNo + vbCritical, "Set Date?")
    If Response = vbNo Then Exit Sub
    On Error GoTo FixMe
    SQLz = "UPDATE UnitExitOption INNER JOIN EXITOptions_WorkTable ON UnitExitOption.UnitID = EXITOptions_WorkTable.UnitID  " & SQLset
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub
Private Sub ActualDateSubmit_Click()
    If Me.EndOrActualDateCBO & "test" = "test" Then
        Exit Sub
    End If
    Dim Response, SQLset As String
    SQLset = "SET EndOrActualDate = '" & Me.EndOrActualDateCBO & "', UserName='" & FAUserName & "', dateposted='" & Now() & "'WHERE OptionTypeID = 2;"
    Response = MsgBox("You're about set the Actual Termination Date for the Lease for all selected units, Are you sure?", vbQuestion + vbYesNo + vbCritical, "Set Date?")
    If Response = vbNo Then Exit Sub
    On Error GoTo FixMe
    SQLz = "UPDATE UnitExitOption INNER JOIN EXITOptions_WorkTable ON UnitExitOption.UnitID = EXITOptions_WorkTable.UnitID  " & SQLset
    '   DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

'Private Sub FeesSubmit_Click()
'    If Me.FeesCBO & "test" = "test" Then
'        Exit Sub
'    End If
'    Dim Response, SQLset As String
'    SQLset = "SET TerminateFee = '" & Me.FeesCBO & "', UserName='" & FAUserName & "', dateposted='" & Now() & "'WHERE OptionTypeID = 2;"
'    Response = MsgBox("You're about set the Term-Fees Expected for all selected units, Are you sure?", vbQuestion + vbYesNo + vbCritical, "Set TermFees?")
'    If Response = vbNo Then Exit Sub
'    On Error GoTo Fixme
'    SQLz = "UPDATE UnitExitOption INNER JOIN EXITOptions_WorkTable ON UnitExitOption.UnitID = EXITOptions_WorkTable.UnitID  " & SQLset
'    DoCmd.SetWarnings False
'    DoCmd.RunSQL (SQLz)
'    DoCmd.SetWarnings True
'Fixme:
'    Me.Form.Refresh
'    Exit Sub
'End Sub
Private Sub ResidualSubmit_Click()
    If Me.ResidualCBO & "test" = "test" Then
        Exit Sub
    End If
    Dim Response, SQLset As String
    SQLset = "SET Residual = " & Me.ResidualCBO & ", UserName='" & FAUserName & "', dateposted='" & Now() & "'WHERE OptionTypeID = 2;"
    Response = MsgBox("You're about set the Residual for all selected units, Are you sure?", vbQuestion + vbYesNo + vbCritical, "Set Residual?")
    If Response = vbNo Then Exit Sub
    On Error GoTo FixMe
    SQLz = "UPDATE UnitExitOption INNER JOIN EXITOptions_WorkTable ON UnitExitOption.UnitID = EXITOptions_WorkTable.UnitID  " & SQLset
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub
Private Sub NotesSubmit_Click()
    If Me.NotesCbo & "test" = "test" Then
        Exit Sub
    End If
    Dim Response, SQLset As String
    SQLset = "SET Notes = '" & Me.NotesCbo & "', UserName='" & FAUserName & "', dateposted='" & Now() & "'WHERE OptionTypeID = 2;"
    Response = MsgBox("You're about set the Notes for all selected units, Are you sure?", vbQuestion + vbYesNo + vbCritical, "Set Notes?")
    If Response = vbNo Then Exit Sub
    On Error GoTo FixMe
    SQLz = "UPDATE UnitExitOption INNER JOIN EXITOptions_WorkTable ON UnitExitOption.UnitID = EXITOptions_WorkTable.UnitID  " & SQLset
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
FixMe:
    Me.Form.Refresh
End Sub

Private Sub Send2Excel_Click()
    Dim XLS As String
    XLS = "|D,HH|I,HH|F,BC|H,$E|J,BC|M,$E|Q,$E|"
    SendFormData2Excel Me, "Units", "Buy Out Info for " & Me.UnitCount & " Units", True, XLS, Me.UnitCount
End Sub

Private Sub UnitID_DblClick(Cancel As Integer)
    Call OpenUnitsForm(Me.UnitID)
    'DoCmd.OpenForm "UnitsForm", , , "UNITID=" & Me.UnitID & ""
    'Forms!UnitsForm.UnitsTabControl.Value = 10
    'Me.UnitNum.SetFocus
End Sub
Public Sub SetModified()
    Me.UserName = FAUserName
    Me.DatePosted = Now()
End Sub
