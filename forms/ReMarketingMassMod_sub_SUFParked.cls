Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public SQLz, SQLset As String
Public Committ As Variant
Dim RST        As DAO.Recordset
Dim i          As Integer
Public Sub ClearBoxes()
    Dim Ccrl   As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Medit" Then
            Me(Ccrl.Name).Value = ""
        End If
    Next Ccrl
End Sub

Private Sub Form_Current()
    On Error Resume Next
    Me.txtCurrent = Me.UnitID
    SQLz = "Select ClientShortname As Client, Schedule, UnitGroup From vw_SixKeys WHERE AssetID=" & Me.UnitID & ";"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    Me.Parent.BuyNo = "Client: " & RST!Client & " Sch|Grp: " & RST!Schedule & "|" & RST!UnitGroup
    RST.close
    Set RST = Nothing
    On Error GoTo 0
End Sub

Private Sub ParkedDate_AfterUpdate()
    If IsDate(Me.ParkedDate) And Me.UnitStatus = "A" Or Me.UnitStatus = "TR" Then SQLz -"Update Units Set UnitStatus='EP' WHERE UnitId=" & Me.UnitID & ";"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
    End If
FixMe:
    Exit Sub
End Sub

Private Sub BankNotifiedSubmit_Click()
    On Error GoTo FixMe
    If Me.BankNotifiedCbo & "test" = "test" Then Call UpdateRMKTtbl("RMKT.BankNotified", Null, , , Me.Parent.Name)
    If IsDate(Me.BankNotifiedCbo) Then Call UpdateRMKTtbl("RMKT.BankNotified", Me.BankNotifiedCbo, , , Me.Parent.Name)
    Me.Form.Refresh
FixMe:
    Exit Sub
End Sub

Private Sub LocParkedSubmit_Click()
    If Me.LocParkedCbo & "test" = "test" Then Exit Sub
    On Error GoTo FixMe
    Call UpdateRUnitTbl("LocationParkedID", Me.LocParkedCbo)
FixMe:
    Exit Sub
End Sub

Private Sub OffLeaseSubmit_Click()
    On Error GoTo FixMe
    If Me.OffLeaseCbo & "test" = "test" Then Call UpdateRUnitTbl("unitoffleasedt", Null, , , Me.Parent.Name)
    If IsDate(Me.OffLeaseCbo) Then Call UpdateRUnitTbl("unitoffleasedt", Me.OffLeaseCbo, , , Me.Parent.Name)
FixMe:
    Exit Sub
End Sub

Private Sub ParkedDateSubmit_Click()
    On Error GoTo FixMe
    If Me.ParkedDateCbo & "test" = "test" Then
       Call UpdateRMKTtbl("RMKT.ParkDT", Null, , , Me.Parent.Name)
       Call UpdateRUnitTbl("Units.UnitStatus", "TR", "WHERE (UnitStatus = 'P' Or NOT UnitStatus Like '*E*')", , Me.Parent.Name)
    End If
    If IsDate(Me.SufSentCBO) Then
       Call UpdateRMKTtbl("RMKT.ParkDT", Me.ParkedDateCbo, , , Me.Parent.Name)
       Call UpdateRUnitTbl("Units.UnitStatus", "EP", "WHERE (UnitStatus = 'TR' Or Units.UnitStatus = 'A')", , Me.Parent.Name)
    End If
    Me.Form.Refresh
FixMe:
    Exit Sub
End Sub

Private Sub SUFReturnedSubmit_Click()
    On Error GoTo FixMe
    If Me.SUFReturnedCbo & "test" = "test" Then
       Call UpdateRMKTtbl("RMKT.[SUF Returned]", Null, , , Me.Parent.Name)
       Call UpdateRUnitTbl("Units.UnitStatus", "ES")
    End If
    If IsDate(Me.SUFReturnedCbo) Then
       Call UpdateRMKTtbl("RMKT.[SUF Returned]", Me.SUFReturnedCbo, , , Me.Parent.Name)
       Call UpdateRUnitTbl("Units.UnitStatus", "EP")
    End If
    Me.Form.Refresh
FixMe:
    Exit Sub
End Sub

Private Sub SUFSentSubmit_Click()
    On Error GoTo FixMe
    If Me.SufSentCBO & "test" = "test" Then
       Call UpdateRMKTtbl("RMKT.[SUF Sent]", Null, , , Me.Parent.Name)
       Call UpdateRUnitTbl("Units.[UnitStatus]", "TR")
    End If
    If IsDate(Me.SufSentCBO) Then
       Call UpdateRMKTtbl("RMKT.[SUF Sent]", Me.SufSentCBO, , , Me.Parent.Name)
       Call UpdateRUnitTbl("Units.[UnitStatus]", "ES")
    End If
    Me.Form.Refresh
FixMe:
    Exit Sub
End Sub

Private Sub UnitStatusSubmit_Click()
    If Me.UnitStatusCbo & "test" = "test" Then Exit Sub
    On Error GoTo FixMe
    Call UpdateRUnitTbl("Units.UnitStatus", UnitStatusCbo, , , Me.Parent.Name)
    Me.Form.Refresh
FixMe:
    Exit Sub
End Sub

Private Sub SaveThisNote_Click()
'Call ConfrimIT("UnitNote", "Single Units Notes")
    If Me.UnitNote & "test" = "test" Then Exit Sub
    If CheckNote(Me.UnitID) = "Already Saved" Then Exit Sub
AskDate:
    Dim NoteSent As String

    NoteSent = InputBox("Provide the DATE for this note." & vbNewLine & _
                      "          Please enter that date, below." & vbNewLine & _
                      "          Today's date is the default." & vbNewLine & vbNewLine & _
                      "          ( You MUST use a date )" _
             , "Note Date:", Date)
    If Not IsDate(NoteSent) Then GoTo AskDate
    SQLz = "Insert Into UnitNotes (UnitRef,uNoteDate,uNote,uEmp)"
    SQLz = SQLz & "Select " & Me.UnitID & ",#" & NoteSent & "#,'" & Me.UnitNote & "','" & FAUserInit & "'"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
FixMe:
    Me.Form.Refresh
End Sub

Private Sub UnitNoteSubmit_Click()
    If Me.UnitNoteCBO & "test" = "test" Or CheckNote = "Already Saved" Then Exit Sub
    Call ConfrimIT("UnitNoteCbo", "Units Notes")
AskDate:
    Dim NoteSent As String

    NoteSent = InputBox("Provide the DATE for this note." & vbNewLine & _
                      "          Please enter that date, below." & vbNewLine & _
                      "          Today's date is the default." & vbNewLine & vbNewLine & _
                      "          ( You MUST use a date )" _
             , "Note Date:", Date)
    If Not IsDate(NoteSent) Then GoTo AskDate
    Me.UnitNote = Me.UnitNoteCBO

    If Me.Parent.Name = "BuyersOrderMassMod" Then
       Set RST = CurrentDb.OpenRecordset("Select UnitId From BuyersOrder_MMod_Work")
    Else
      Set RST = CurrentDb.OpenRecordset("Select UnitId From RMKT_MMod_Work")
    End If
    RST.MoveLast
    RST.MoveFirst
    DoCmd.SetWarnings False
    For i = 1 To RST.RecordCount
        SQLz = "Insert Into UnitNotes (UnitRef,uNoteDate,uNote,uEmp)"
        SQLz = SQLz & "Select " & RST!UnitID & ",#" & NoteSent & "#,'" & Me.UnitNoteCBO & "','" & FAUserInit & "'"
        DoCmd.RunSQL (SQLz)
        RST.MoveNext
    Next i
FixMe:
    DoCmd.SetWarnings True
    RST.close
    Set RST = Nothing
    Me.Form.Refresh
End Sub
Public Function CheckNote(Optional UnitIDx As Long) As String
    SQLz = "SELECT RMKT_MMod_Work.UnitID, UnitNotes.uNoteDate, UnitNotes.uNote "
    SQLz = SQLz & "FROM UnitNotes INNER JOIN RMKT_MMod_Work ON UnitNotes.UnitRef = RMKT_MMod_Work.UnitID "
    If UnitIDx <> 0 Then SQLz = SQLz & "Where UnitId =" & UnitID & ";"
    Set RST = CurrentDb.OpenRecordset(SQLz)
    On Error GoTo FixMe
    RST.MoveLast
    RST.MoveFirst
    On Error GoTo 0
    For i = 1 To RST.RecordCount
        If Me.UnitNote = RST!uNote Then
            MsgBox "This Note has already been saved" & vbNewLine & "No action has been taken", vbCritical, "You're repeating yourself"
            CheckNote = "Already Saved"
            GoTo FixMe
        End If
        RST.MoveNext
    Next i
FixMe:
    RST.close
    Set RST = Nothing
End Function
Private Sub MileageSubmit_Click()
   Exit Sub
End Sub
Private Sub HoursSubmit_Click()
    Exit Sub
End Sub

Public Sub ConfrimIT(ByVal FieldBoX As String, ByVal TragetX As String)
    If Me(FieldBoX).Value & "test" = "test" Then
        Committ = MsgBox("You're about set the " & TragetX & " to Null, for ALL units listed." & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Update-Change confrim " & TragetX & "?")
    Else
        Committ = MsgBox("You're about set the " & TragetX & " to " & Me(FieldBoX).Value & vbNewLine & "for ALL units listed." & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Change " & TragetX & "?")
    End If
    If Committ = vbNo Then Exit Sub
End Sub

Private Sub Send2Excel_Click()
    Dim XLS As String
    XLS = "|D,HH|I,HH|F,BC|H,$E|J,BC|M,$E|Q,$E|"
    SendFormData2Excel Me, "Units", "Buy Out Info for " & Me.UnitCount & " Units", True, XLS, Me.UnitCount
End Sub

Private Sub UnitID_DblClick(Cancel As Integer)
    Call OpenUnitsForm(Me.UnitID)
    'DoCmd.OpenForm "UnitsForm", , , "UNITID=" & Me.UnitID & ""
    Forms!Unitsform.UnitsTabControl.Value = 6
    Me.UnitNum.SetFocus
End Sub
