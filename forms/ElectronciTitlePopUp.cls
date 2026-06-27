Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub CloseMe_Click()
    DoCmd.close acForm, Me.Form.Name
    Exit Sub
FixMe:
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub ExcelExport_Click()
    If Me.Commission & "test" = "test" Or Me.Signed & "test" = "test" Then
        MsgBox "You must to set a Commission Expires AND the Signed Date", vbCritical, "Why are you so silly?"
        Exit Sub
    End If
Dim SQLz As String
SQLz = "SELECT TOP 100"
SQLz = SQLz & " SchGrp"
SQLz = SQLz & ",Lastsent"
SQLz = SQLz & ",MYear"
SQLz = SQLz & ",Make"
SQLz = SQLz & ",VIN"
SQLz = SQLz & ",#" & Format(Me.Signed, "mm/dd/yyyy") & "# AS  Signed "
SQLz = SQLz & ",tAssetID as [FA AssetID]"
SQLz = SQLz & " FROM ElectronicTitleRequest"
SQLz = SQLz & " WHERE ISNULL(Sentdate) "
SQLz = SQLz & " ORDER BY Lastsent, VIN;"
XLS = "|A,W15|B,TC|B,W10|C,TC|D,TC|D,TC|F,TC|G,BC|"
SendReport2Excel SQLz, "Paper Request", "Paper Title Requested: " & Format(Date, "mm-dd-yy"), True, XLS

End Sub

Private Sub RunBatch_Click()
    If Me.Commission & "test" = "test" Or Me.Signed & "test" = "test" Then
        MsgBox "You must to set a Commission Expires AND the Signed Date", vbCritical, "Why are you so silly?"
        Exit Sub
    End If

    DoCmd.OpenReport "ElectronicTitleRequestOK", acViewPreview

End Sub

Private Sub RunHistory_Click()
    If Me.Commission & "test" = "test" Or Me.Signed & "test" = "test" Then
        MsgBox "You must to set a Commission Expires AND the Signed Date", vbCritical, "Why are you so silly?"
        Exit Sub
    End If
    Dim Committ As String
    Dim SQLx As String
    Dim SQLz As String
    Committ = MsgBox("You're about to mark this batch as 'Done'" & vbNewLine & _
        "and Set 'Sent to State for Paper Title' Log." & vbNewLine & _
        "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Set Date?")
    If Committ = vbNo Then Exit Sub
    Dim RST As DAO.Recordset
    SQLz = "SELECT TOP 100 tAssetID"
    SQLz = SQLz & " FROM ElectronicTitleRequest"
    SQLz = SQLz & " WHERE (((SentDate) Is Null))"
    SQLz = SQLz & " ORDER BY Lastsent, VIN;"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    On Error GoTo FixMe
    RST.MoveLast
    RST.MoveFirst
    If RST.RecordCount > 0 Then
        MessageUser "Open", "Marking Batch / Setting Log file", "This will take a moment or two"
        For i = 1 To RST.RecordCount
            SQLx = "Update ElectronicTitleRequest Set SentDate = #" & CDate(Me.Signed) & "# Where tAssetID=" & RST!tAssetID & ";"
            SQLz = "INSERT INTO UnitTitleLog ( UnitRef, EntryReason,  UnitTitleLogDate,  EntryTypeId,  EntryReasonId,  ExpectedReturn,  UserID )"
            SQLz = SQLz & "SELECT " & RST!tAssetID & ", 'Sent to State for Paper Title', #" & CDate(Me.Signed) & "#, 2, 49, True, '" & FAUserName & "';"
            DoCmd.SetWarnings False
            DoCmd.RunSQL (SQLx)
            DoCmd.RunSQL (SQLz)
            DoCmd.SetWarnings True
            RST.MoveNext
        Next i
    End If
FixMe:
    On Error GoTo 0
    RST.Clone
    Set RST = Nothing
    DoCmd.SetWarnings True
    MessageUser "Close"
    Me.Commission = ""
    Me.Signed = ""
    Me.BatchList.Requery
End Sub
