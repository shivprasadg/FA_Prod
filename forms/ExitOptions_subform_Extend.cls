Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public SQLz, SQLx    As String
Public EndDatex As Date
Public Sub ClearBoxes()
    Dim Ccrl   As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Medit" Then
            Me(Ccrl.Name).Value = ""
        End If
    Next Ccrl
End Sub


Private Sub AddExtend_Click()
    DoCmd.RunCommand acCmdSaveRecord
    If Not IsDate(Me.StartOrPurposedDate) Or Not IsDate(Me.EndOrActualDate) Then
        MsgBox "The Current ExtendIT Option Dosen't Have A Complete Date Range Set" & vbNewLine & "Please Set Both Start and End Dates", vbCritical, "Need Dates"
        Exit Sub
    End If
    Dim Response, SQLset, ExtType, ExtSubType, Msgx As String
    Dim DateEndX As Date
    Dim iOption As Integer
    Dim DateStartX, TermX, Rentx, FeeX, ResiX, ResiOldX, ReplaceEndX, ReplaceEndNoteX As Variant
    Response = MsgBox("You're about to Add a NEW Extension to an Existing Extended Lease..." & vbNewLine & "Are you sure?", vbQuestion + vbYesNo + vbCritical, "Extend an Extension.")
    If Response = vbNo Then Exit Sub
    Response = MsgBox("Is this a M-T-M extension?" & vbNewLine & "Yes or No", vbQuestion + vbYesNo, "What Type of Extension.")
    iOption = Nz(DMax("OptionTypeID", "UnitExitOption", "OptionTypeID>2 AND UnitID=" & Me.UnitID & ""), 2) + 1
    If Response = vbYes Then
        ExtType = "HdOvr"
        ExtSubType = "M-T-M"
    Else
        ExtType = "Extend"
        ExtSubType = "Term"
    End If
    Response = MsgBox("Will this 'new' extension start at the end of the current extension (" & Me.EndOrActualDate & ")?" & vbNewLine & "Yes or No", vbQuestion + vbYesNo, "What Type of Extension.")
    If Response = vbYes Then
        DateStartX = Me.EndOrActualDate + 1
        GoTo SetTerm
    End If
SetStart:
    DateStartX = InputBox("What is the Start Date?" & vbNewLine & "please keep correct date format mm/dd/yyyy", "Date this Extened Lease is to start.", Me.EndOrActualDate + 1)
    If IsDate(DateStartX) Then
        If CDate(DateStartX) < Me.EndOrActualDate Then
            Response = MsgBox(DateStartX & " is before the end of the current extension(" & Me.EndOrActualDate & ")" & vbNewLine & "Do you want to replace the existing extension's end date with " & CDate(DateStartX) - 1 & "?", vbQuestion + vbYesNo + vbCritical, "Extend an Extension.")
            If Response = vbNo Then
                GoTo SetStart
            Else
                ReplaceEndX = CDate(DateStartX) - 1
                ReplaceEndNoteX = FAUserName & " changed the prev extension's EOL date from " & Me.EndOrActualDate & " to " & CDate(DateStartX) - 1 & " to match new lease extension parameters"
            End If
        End If
    Else
        GoTo SetStart
    End If
SetTerm:
    TermX = InputBox("How Many Months to Extend This Lease?" & vbNewLine & "numeric values only please", "Term in months.", 12)
    If IsNumeric(TermX) Then Else GoTo SetTerm
SetRent:
    Rentx = InputBox("What is the Rent for this unit?" & vbNewLine & "Dollars and Cents... you can enter 0.", "Rent to Charge...", Format(Nz(Me.Rent, 0), "$#,#0.#0"))
    If IsNumeric(Rentx) Then Else GoTo SetRent
SetFees:
    FeeX = InputBox("Please set the Service Fees for this unit." & vbNewLine & "Dollars and Cents... you can enter 0", "Service Fee to Charge.", Format(Nz(Me.Fees, 0), "$#,#0.#0"))
    If IsNumeric(FeeX) Then Else GoTo SetFees
SetResi:
    ResiX = Me.Residual
    'ResiX = InputBox("Please set the Residual Value for this Unit." & vbNewLine & "The previous value was " & Format(Me.Residual, "$#,#0.#0") & vbNewLine & " Dollars and Cents... you can enter 0", "Set Residual", Format(Nz(Me.Residual, 0), "$#,#0.#0"))
    'If IsNumeric(ResiX) Then Else GoTo SetResi
    Msgx = "You want to add a " & ExtType & " to this Unit " & vbNewLine & _
           "with a start date of:    " & DateStartX & vbNewLine & _
           "with a term of:            " & TermX & " months " & vbNewLine & _
           "lease ending on:         " & DateAdd("M", TermX, DateStartX) - 1 & vbNewLine & _
           "the rent is set at:          " & Rentx & vbNewLine & _
           "with a fee of:                " & FeeX & vbNewLine
    '       "the residual is:              " & ResiX & vbNewLine & ""
    If IsDate(ReplaceEndX) Then Msgx = Msgx & "Replacing the existing extension's end date to " & CDate(ReplaceEndX) & vbNewLine & ""
    Msgx = Msgx & vbNewLine & "Please confrim, Yes or No, (No will cancel)"
    Response = MsgBox(Msgx, vbQuestion + vbYesNo, "Confrim Extend an Extension")
    If Response = vbNo Then
        Exit Sub
    Else
        If IsDate(ReplaceEndX) Then
            Me.EndOrActualDate = ReplaceEndX
            Me.Notes = Me.Notes & " <AC> " & ReplaceEndNoteX
            Me.Term = DateDiff("M", Me.StartOrPurposedDate, Me.EndOrActualDate)
        End If
        Me.OptionTypeID = iOption
        Me.UIDX = Me.UnitID & "-" & iOption
        SQLz = "INSERT INTO UnitExitOption ( UserName, UnitID, OptionTypeID, OptionTypeDesc, StartOrPurposedDate, EndOrActualDate, TermMonths, Rent, Fees, Residual, DatePosted, OptionSubType, Notes, UIDX ) " & _
               "SELECT '" & FAUserName & "', " & Me.UnitID & ", " & iOption & ", '" & ExtType & "', '" & DateStartX & "', '" & DateAdd("M", TermX, DateStartX) - 1 & "', " & TermX & ", '" & Rentx & "', '" & FeeX & "', '" & ResiX & "','" & Now() & "', '" & ExtSubType & "','This is the 2nd Extension','" & Me.UnitID & "-3';"
        Me.Form.Requery
        DoCmd.SetWarnings False
        DoCmd.RunSQL ("UPDATE EXITOptions_WorkTable SET UID3='" & Me.UnitID & "-" & iOption & "' WHERE UnitID=" & Me.UnitID & ";")
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
        Me.Form.Requery
    End If


End Sub

Private Sub AssigneeNotified_Change()
    Call SetModified
End Sub

Private Sub AssigneeReplied_Change()
    Call SetModified
End Sub

Private Sub DeleteAll_Click()
    Dim Response, SQLset As String
    Response = MsgBox("You're about to Delete all Extend-IT options for all of the listed units " & vbNewLine & Me.OptionTypeDesc & " option entry , " & vbNewLine & "Are you sure?", vbQuestion + vbYesNo + vbCritical, "What? Delete!")
    If Response = vbNo Then Exit Sub
    SQLset = " WHERE OptionTypeID = 3;"
    On Error GoTo FixMe
    SQLz = "DELETE UnitExitOption.* From UnitExitOption INNER JOIN EXITOptions_WorkTable ON UnitExitOption.UnitID = EXITOptions_WorkTable.UnitID  " & SQLset
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    SQLz = "Delete  * From EXITOptions_WorkTable"
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Me.Form.Requery
    Me.Parent.Units_List.Requery
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub DeleteMe_Click()
    Dim Response As String
    Response = MsgBox("You're about to Delete Unit# " & Me.UnitUnitNum & "'s " & vbNewLine & Me.OptionTypeDesc & " option entry , " & vbNewLine & "Are you sure?", vbQuestion + vbYesNo + vbCritical, "What? Delete!")
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

Private Sub EndOrActualDate_Change()
    If Not IsDate(Me.EndOrActualDate) Then Exit Sub
    Call SetModified
    On Error GoTo FixMe
    DoCmd.RunCommand acCmdSaveRecord
    Me.Parent.Units_List.Requery
FixMe:
End Sub



Private Sub Fees_Change()
    Call SetModified
End Sub

Private Sub Form_Current()
On Error Resume Next
    Me.txtCurrent = Me.UnitID
On Error GoTo 0
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


Private Sub ModifyExch_Click()
    DoCmd.RunCommand acCmdSaveRecord
    Form_EXITOptions.RemoteSelectSingleUnit (Me.UnitID)
    Form_EXITOptions.ExchangeSelected_Click
End Sub

Private Sub ModifyTERM_Click()
    DoCmd.RunCommand acCmdSaveRecord
    Form_EXITOptions.RemoteSelectSingleUnit (Me.UnitID)
    Form_EXITOptions.TerminateSelected_Click
End Sub

Private Sub Notes_Change()
    Call SetModified
End Sub

Public Sub OptionSubType_Click()
    Dim RST    As DAO.Recordset
    Me.Term = ""
    SQLz = "Select * From SchGrp WHERE sGrpID = " & Me.groupID & ""
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    RST.MoveLast
    DateEndX = DateAdd("m", RST!LeaseTermPrimary, RST!BLCD)
    If Me.OptionSubType Like "*M-T-M*" Then
        Me.Rent = CCur(Nz(RST!MTMRent, 0))
        Me.Term = 3
    Else
        Me.Rent = CCur(Nz(RST!LeaseExtOption1Rent, 0))
        Me.Term = CInt(Nz(RST!LeaseExtOption1Term, 12))
    End If
    Me.Fees = CCur(Nz(RST!SGrpFSLPmt, 0))
    Me.StartOrPurposedDate = DateEndX
    If Me.Term > 0 Then Me.EndOrActualDate = DateAdd("m", Me.Term, DateEndX) - 1
    RST.close
    Set RST = Nothing
    Call SetModified
    DoCmd.RunCommand acCmdSaveRecord
    Me.Parent.Units_List.Requery
End Sub

Private Sub OptionTypeDesc_Click()
       SQLx = "SELECT OptionSubType, OptionSubTypeDesc, ExitOptionType FROM LookUp_ExitOptionTypes " & _
              "WHERE ExitOptionType='" & OptionTypeDesc & "' Order By OptionSubType ;"
    Me.OptionSubTypeCbo.RowSource = SQLx
    If OptionTypeDesc = "HdOvr" Then OptionSubType = "M-T-M" Else OptionSubType = ""
End Sub

Private Sub OptionTypeDescCBO_Click()
    SQLx = "SELECT OptionSubType, OptionSubTypeDesc, ExitOptionType FROM LookUp_ExitOptionTypes " & _
              "WHERE ExitOptionType='" & OptionTypeDescCBO & "' Order By OptionSubType ;"
    Me.OptionSubTypeCbo.RowSource = SQLx
    If OptionTypeDescCBO = "HdOvr" Then OptionSubTypeCbo = "M-T-M" Else OptionSubTypeCbo = ""

End Sub

Public Sub OptionTypeDescSubmit_Click()
       If Me.OptionTypeDescCBO & "test" = "test" Then
        Exit Sub
    End If
     Dim Response, SQLset As String
    SQLset = "SET OptionTypeDesc = '" & Me.OptionTypeDescCBO & "', UserName='" & FAUserName & "', dateposted='" & Now() & "' WHERE OptionTypeID = " & Me.OptionTypeID & ";"
    Response = MsgBox("You're about set the Exteded Lease Type for all selected units, Are you sure?", vbQuestion + vbYesNo + vbCritical, "Set Type?")
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
Private Sub OptionSubTypeSubmit_Click()
    If Me.OptionSubTypeCbo & "test" = "test" Then
        Exit Sub
    End If
    Dim Response, SQLset As String
    SQLset = "SET OptionSubType = '" & Me.OptionSubTypeCbo & "', UserName='" & FAUserName & "', dateposted='" & Now() & "' WHERE OptionTypeID = " & Me.OptionTypeID & ";"
    Response = MsgBox("You're about set the Exteded Lease SUB Type for all selected units, Are you sure?", vbQuestion + vbYesNo + vbCritical, "Set Type?")
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

Private Sub Rent_Change()
    Call SetModified
End Sub

Private Sub RentSubmit_Click()
    If Me.RentCBO & "test" = "test" Then
        Exit Sub
    End If
    Dim Response, SQLset As String
    SQLset = "SET Rent= '" & Me.RentCBO & "', UserName='" & FAUserName & "', dateposted='" & Now() & "'WHERE OptionTypeID = 3;"
    Response = MsgBox("You're about set the Rent for all selected units, Are you sure?", vbQuestion + vbYesNo + vbCritical, "Set Extended Rent?")
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

Private Sub Residual_Change()
    Call SetModified
End Sub

Private Sub SaveUpdateMe_Click()
    Call MakeExitOptionLocal(False, True)
    Form_EXITOptions.SetGapLapData (3)
    Me.Form.Refresh
End Sub

Private Sub StartOrPurposedDate_BeforeUpdate(Cancel As Integer)
    DateEndX = DateAdd("m", DLookup("LeaseTermPrimary", "SchGrp", "sGrpID = " & Me.groupID), DLookup("BLCD", "SchGrp", "sGrpID = " & Me.groupID))
    If IsDate(DateEndX) Then
        If Me.StartOrPurposedDate < DateEndX Then
            MsgBox "The date you selected: " & StartOrPurposedDate & " is BEFORE" & vbNewLine & _
                   "the end date of the ACTIVE lease: " & DateEndX - 1 & vbNewLine & vbNewLine & _
                   "The Start Date will be set to the previous date.", vbCritical, "Extended Lease Date Conflict"
            Me.Undo
        End If
        If Me.StartOrPurposedDate > DateEndX Then
            MsgBox "The date you selected: " & StartOrPurposedDate & " will create a LEASE GAP" & vbNewLine & _
                   "The date you selected is AFTER the end of ACTIVE lease: " & DateEndX - 1 & vbNewLine & vbNewLine & _
                   "The Start Date will be set to the previous date.", vbCritical, "Extended Lease Date Conflict"
            Me.Undo
        End If
    End If
End Sub
Private Sub StartOrPurposedDateCBO_AfterUpdate()
    If Me.TermCBO & "test" <> "test" Then Me.EndOrActualDateCBO = DateAdd("M", Me.TermCBO, Me.StartOrPurposedDateCBO) - 1
    If Me.TermCBO & "test" = "test" And IsDate(Me.EndOrActualDateCBO) Then Me.TermCBO = DateDiff("M", Me.StartOrPurposedDateCBO, Me.EndOrActualDateCBO)
    DoCmd.RunCommand acCmdSaveRecord
End Sub
Private Sub EndOrActualDateCBO_AfterUpdate()
If Me.TermCBO & "test" <> "test" Then Me.StartOrPurposedDateCBO = DateAdd("M", (Me.TermCBO * -1), Me.EndOrActualDateCBO) - 1
If Me.TermCBO & "test" = "test" And IsDate(Me.StartOrPurposedDateCBO) Then Me.TermCBO = DateDiff("M", Me.StartOrPurposedDateCBO, Me.EndOrActualDateCBO)
If Not IsDate(Me.StartOrPurposedDateCBO) Then
MsgBox "Please insert a Start Date", vbCritical, AtlaasVersion
Me.StartOrPurposedDateCBO.SetFocus
End If
End Sub
Private Sub Term_Change()
    Call SetModified
    DoCmd.RunCommand acCmdSaveRecord
    Me.Parent.Units_List.Requery
End Sub

Private Sub TermCBO_Click()
    If Not IsDate(StartOrPurposedDateCBO) Then
        MsgBox "Please make sure to input a Start Date", vbInformation, atlassversion
        Me.StartOrPurposedDateCBO.SetFocus
        Exit Sub
    End If
    If IsDate(Me.StartOrPurposedDateCBO) Then Me.EndOrActualDateCBO = DateAdd("M", Me.TermCBO, Me.StartOrPurposedDateCBO) - 1
End Sub
Private Sub Term_Click()
    If Me.Term & "test" = "test" Then
        Exit Sub
    End If
    If IsDate(Me.StartOrPurposedDate) Then Me.EndOrActualDate = DateAdd("M", Me.Term, Me.StartOrPurposedDate) - 1

End Sub
Private Sub StartOrPurposedDate_AfterUpdate()
    If Me.Term & "test" <> "test" Then Me.EndOrActualDate = DateAdd("M", Me.Term, Me.StartOrPurposedDate) - 1
    If Me.Term & "test" = "test" And IsDate(Me.EndOrActualDate) Then Me.TermCBO = DateDiff("M", Me.StartOrPurposedDate, Me.EndOrActualDate)
End Sub
Private Sub TermSubmit_Click()
    If Me.TermCBO & "test" = "test" Then
        Exit Sub
    End If
    If Me.StartOrPurposedDateCBO & "test" = "test" Then
        MsgBox "Please enter a Start Date", vbCritical, atlassversion
        Me.StartOrPurposedDateCBO.SetFocus
        Exit Sub
    End If
    Dim Response, SQLset As String
    SQLset = "SET TermMonths = " & Me.TermCBO.Column(0) & ", "
    If IsDate(Me.StartOrPurposedDateCBO) Then SQLset = SQLset & "StartOrPurposedDate= '" & StartOrPurposedDateCBO & "', EndOrActualDate='" & DateAdd("M", Me.TermCBO, Me.StartOrPurposedDateCBO) - 1 & "', "
    SQLset = SQLset & "UserName='" & FAUserName & "', dateposted='" & Now() & "'WHERE OptionTypeID = 3;"
    Response = MsgBox("You're about set the Term in Month for all selected units, Are you sure?", vbQuestion + vbYesNo + vbCritical, "Set Extended Term?")
    If Response = vbNo Then Exit Sub
    On Error GoTo FixMe
    SQLz = "UPDATE UnitExitOption INNER JOIN EXITOptions_WorkTable ON UnitExitOption.UnitID = EXITOptions_WorkTable.UnitID  " & SQLset
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Call SetModified
FixMe:
    Me.Form.Refresh
    Exit Sub

End Sub
Private Sub PurposedSubmit_Click()
    If Me.StartOrPurposedDateCBO & "test" = "test" Then
        Exit Sub
    End If
    Dim Response, SQLset As String
    SQLset = "SET StartOrPurposedDate = '" & Me.StartOrPurposedDateCBO & "', UserName='" & FAUserName & "', dateposted='" & Now() & "'WHERE OptionTypeID = 3;"
    Response = MsgBox("You're about set the Extended Start Date for all selected units, Are you sure?", vbQuestion + vbYesNo + vbCritical, "Set Date?")
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
    If Me.TermCBO & "test" = "test" Then
        MsgBox "You must set the number of months", vbCritical, atlassversion
        Me.TermCBO.SetFocus
        Exit Sub
    End If
    Dim Response, SQLset As String
    'SQLset = "SET EndOrActualDate = '" & Me.EndOrActualDateCBO & "', UserName='" & FAUserName & "', dateposted='" & Now() & "'WHERE OptionTypeID = 3;"
    Response = MsgBox("You're about set the End Date for the Extended Lease for all selected units, Are you sure?", vbQuestion + vbYesNo + vbCritical, "Set Date?")
    If Response = vbNo Then Exit Sub
    On Error GoTo FixMe
    Call TermSubmit_Click
    'Sqlz = "UPDATE UnitExitOption INNER JOIN EXITOptions_WorkTable ON UnitExitOption.UnitID = EXITOptions_WorkTable.UnitID  " & SQLset
    'DoCmd.SetWarnings False
    'DoCmd.RunSQL (Sqlz)
    'DoCmd.SetWarnings True
FixMe:
    'Me.Form.Refresh
    Exit Sub
End Sub
'Private Sub FinalRentDateSubmit_Click()
'    If Me.FinalRentDateCBO & "test" = "test" Then
'        Exit Sub
'    End If
'    Dim Response, SQLset As String
'    SQLset = "SET FinalRentDate = '" & Me.FinalRentDateCBO & "', UserName='" & FAUserName & "', dateposted='" & Now() & "'WHERE OptionTypeID = 3;"
'    Response = MsgBox("You're about set the Final Rent Date for all selected units, Are you sure?", vbQuestion + vbYesNo + vbCritical, "Set Final Rent Date?")
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
    SQLset = "SET Residual = " & Me.ResidualCBO & ", UserName='" & FAUserName & "', dateposted='" & Now() & "'WHERE OptionTypeID = 3;"
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
    SQLset = "SET Notes = '" & Me.NotesCbo & "', UserName='" & FAUserName & "', dateposted='" & Now() & "'WHERE OptionTypeID = 3;"
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

'Private Sub SentSubmit_Click()
'    If Me.AssigneeNotifiedCBO & "test" = "test" Then
'        Exit Sub
'    End If
'    Dim Response, SQLset As String
'    SQLset = "SET AssigneeNotified = '" & Me.AssigneeNotifiedCBO & "', UserName='" & FAUserName & "', dateposted='" & Now() & "'WHERE OptionTypeID = 3;"
'    Response = MsgBox("You're about set the Assignee Notified Date for all selected units, Are you sure?", vbQuestion + vbYesNo + vbCritical, "Set Assignee Notified?")
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
'Private Sub RecvSubmit_Click()
'    If Me.AssigneeRepliedCBO & "test" = "test" Then
'        Exit Sub
'    End If
'    Dim Response, SQLset As String
'    SQLset = "SET AssigneeReplied = '" & Me.AssigneeRepliedCBO & "', UserName='" & FAUserName & "', dateposted='" & Now() & "'WHERE OptionTypeID = 3;"
'    Response = MsgBox("You're about set the Assignee Replied Date for all selected units, Are you sure?", vbQuestion + vbYesNo + vbCritical, "Set Assignee Replied?")
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
Private Sub Send2Excel_Click()
    Dim XLS As String
    XLS = "|D,HH|I,HH|F,BC|H,$E|J,BC|M,$E|Q,$E|"
    SendFormData2Excel Me, "Units", "Buy Out Info for " & Me.UnitCount & " Units", True, XLS, Me.UnitCount
End Sub

Private Sub UnitID_DblClick(Cancel As Integer)
    Call OpenUnitsForm(Me.UnitID)
    'DoCmd.OpenForm "UnitsForm", , , "UNITID=" & Me.UnitID & ""
    Forms!Unitsform.UnitsTabControl.Value = 8
    'Me.UnitNum.SetFocus
End Sub
Public Sub SetModified()
    Me.UserName = FAUserName
    Me.DatePosted = Now
End Sub
