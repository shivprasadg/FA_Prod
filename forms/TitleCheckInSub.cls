Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Dim SQLz, MissingX As String

Private Sub EntryReason_Click()
Me.EntryReasonId = Me.EntryReason.Column(2)
End Sub

Private Sub Form_Load()
    Me.RecordSource = ""
End Sub

Private Sub LogTitleIn_Click()
    If Nz(DLookup("EntryTypeID", "UnitTitleHistoryLastEntry", "UnitRef=" & Me.UnitID), 0) <> 2 Then
        MsgBox "You're Trying to Log IN a Title That is Already 'IN'" & vbNewLine & "No action was taken", vbCritical, "You got peanut butter in my chocolate"
        Exit Sub
    End If
    Call CheckFields
    If MissingX <> "You're Missing:" Then
        MsgBox MissingX, vbCritical, "Data missing"
        Exit Sub
    End If
    SQLz = "INSERT INTO UnitTitleLog ( Unitref, TitleNumber, TitleSt, EntryReason, EntryReasonID, UnitTitleLogDate, EntryTypeId, UserID) "
    SQLz = SQLz & "SELECT " & Me.UnitID.Value & ", '" & Me.UnitTitleNumber.Value & "'"
    SQLz = SQLz & ", '" & Me.TitledState & "', '" & Me.EntryReason.Column(0) & "', " & Me.EntryReasonId & " "
    SQLz = SQLz & ", '" & Me.UnitTitleLogDate & "', 1,'" & FaUserName & "';"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Call UpdateUnitTitleState(Me.UnitID)
    Call UpdateUnitTitleNum(Me.UnitID)
    MsgBox "Title # " & Me.UnitTitleNumber & " has been checked IN" & vbNewLine & " The Form will Close", , " Checked In"
    Me.Parent.seekvin.SetFocus
    Me.Parent.TitleLogList.Requery
    Me.Form.Visible = False
End Sub

Private Sub LogTitleOut_Click()
    If Nz(DLookup("EntryTypeID", "UnitTitleHistoryLastEntry", "UnitRef=" & Me.UnitID), 0) <> 1 Then
        MsgBox "You're Trying to Log OUT a Title That is Already 'OUT'" & vbNewLine & "No action was taken", vbCritical, "You're tearing us apart"
        Exit Sub
    End If
    Call CheckFields
    If MissingX <> "You're Missing:" Then
        MsgBox MissingX, vbCritical, "Data missing"
        Exit Sub
    End If
    SQLz = "INSERT INTO UnitTitleLog ( Unitref, TitleNumber, TitleSt, EntryReason, EntryReasonID, UnitTitleLogDate, EntryTypeId, UserID ) "
    SQLz = SQLz & "SELECT " & Me.UnitID.Value & ", '" & Me.UnitTitleNumber.Value & "', '" & Me.TitledState & "'"
    SQLz = SQLz & ", '" & Me.EntryReason & "', " & Me.EntryReasonId & " "
    SQLz = SQLz & ", '" & Me.UnitTitleLogDate & "', 2 ,'" & FaUserName & "';"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Call UpdateUnitTitleState(Me.UnitID)
    Call UpdateUnitTitleNum(Me.UnitID)
    MsgBox "Title # " & Me.UnitTitleNumber & " has been checked OUT" & vbNewLine & " The Form will Close", , " Checked OUT"
    Me.Parent.seekvin.SetFocus
    Me.Parent.TitleLogList.Requery
    Me.Form.Visible = False
End Sub
Public Sub CheckFields()
    MissingX = "You're Missing:"
    If Me.EntryReason & "test" = "test" Then MissingX = MissingX & vbNewLine & "  Reason"
    If Me.UnitTitleNumber & "test" = "test" Then MissingX = MissingX & vbNewLine & "  Title Number"
    If Me.TitledState & "test" = "test" Then MissingX = MissingX & vbNewLine & "  Titled State"
    If Me.UnitYr & "test" = "test" Then MissingX = MissingX & vbNewLine & "  Unit Year"
    If MissingX <> "You're Missing:" Then MissingX = MissingX & vbNewLine & "Exiting... No actions were logged"
End Sub
