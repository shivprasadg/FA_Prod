Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public SQLz As String
Public RST    As DAO.Recordset
Dim Committ As String
Public Sub ClearBoxes()
    Dim Ccrl As Control
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

Private Sub MakeBankPayOffEmail_Click()
    Committ = MsgBox("You're about send the 'Bank Pay-Off Invoice Request' Emails to Ysis." & vbNewLine & "Do you agree?", vbYesNo, "Think! " & TragetX & "?")
    If Committ = vbNo Then Exit Sub
    Call SendRequestPayOffEmail
End Sub
Function makeUnitList()
    UnitIDx = ""
    Set RST = CurrentDb.OpenRecordset("Select UnitID From RMKT_MMod_Work", dbOpenSnapshot)
    On Error Resume Next
    RST.MoveLast
    RST.MoveFirst
    On Error GoTo 0
    If RST.RecordCount > 0 Then
        For i = 1 To RST.RecordCount
            UnitIDx = UnitIDx & RST!UnitID & ","
        Next i
    End If
    RST.close
    Set RST = Nothing
    UnitIDx = Left(UnitIDx, Len(UnitIDx) - 1)
End Function


Private Sub UpdateRMKTField(FieldName As String, controlValue As Variant)
    Dim v As Variant

    v = controlValue
    SQLz = "UPDATE RMKT " & _
           "INNER JOIN BuyersOrder_MMod_Work " & _
           "ON RMKT.UnitRef = BuyersOrder_MMod_Work.UnitID " & vbCrLf

    Select Case True
        Case IsNull(v) Or v = "":   SQLz = SQLz & "SET RMKT." & FieldName & " = NULL;"                 ' NULL
        Case IsDate(v):             SQLz = SQLz & "SET RMKT." & FieldName & " = #" & CDate(v) & "#;"   ' Date
        Case IsNumeric(v):          SQLz = SQLz & "SET RMKT." & FieldName & " = " & v & ";"            ' Integer, Long, Double, Currency
        Case Else:                  SQLz = SQLz & "SET RMKT." & FieldName & " = '" & CStr(v) & "';"    ' String
    End Select

    If IsNull(v) Or v = "" Then
        Committ = MsgBox("You're about set the " & FieldName & " to NULL, for ALL units listed." & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Clear " & FieldName & "?")
    Else
        Committ = MsgBox("You're about set the " & FieldName & " to " & controlValue & ", for ALL units listed." & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Change " & FieldName & "?")
    End If
    If Committ = vbNo Then Exit Sub
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    On Error GoTo 0
    If Err.Number <> 0 Then MsgBox "UpdateRMKTField SQL Error: " & Err.Description
    Me.Refresh
End Sub

Private Sub PayOffReceivedSubmit_Click()
    UpdateRMKTField "PayOffReceived", Me.PayOffReceivedCBO
End Sub

Private Sub PayOffRequestedSubmit_Click()
    UpdateRMKTField "PayOffRequested", Me.PayOffRequestedCBO
End Sub

Private Sub UnitID_DblClick(Cancel As Integer)
    Call OpenUnitsForm(Me.UnitID)
    Forms!Unitsform.UnitsTabControl.Value = 6
    Me.UnitNum.SetFocus
End Sub
