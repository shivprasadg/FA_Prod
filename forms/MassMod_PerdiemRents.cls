Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz As String
Dim ModMade As Boolean

Private Sub CloseMe_Click()
    DoCmd.OpenForm "FAMainMenu"
    If ModMade = -1 Then Call CheckATFChanges(Me.GroupID, "Rent Change")
    DoCmd.close acForm, Me.Form.Name
End Sub


Private Sub Form_Current()
    On Error GoTo whoops
    Me.txtCurrent = Me.UnitID
whoops:
    Exit Sub
End Sub

Public Sub ClearBoxes()
    Dim Ccrl As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Medit" Then
            Me(Ccrl.Name).Value = ""
        End If
    Next Ccrl
    ' Me.DrawDownCbo.Requery
    ' Me.DrawDownId.Requery
End Sub

Private Sub Form_Load()
If Nz(Me.OpenArgs, "") <> "" Then
'Me.Form.RecordSource = "Select * From Units Where GroupRef=" & SplitString(Me.OpenArgs, "?", 0) & ";"
Me.Form.RecordSource = "SELECT * From EditPerdiem WHERE GroupREF=" & SplitString(Me.OpenArgs, "?", 0) & ";"
Me.SchNo = SplitString(Me.OpenArgs, "?", 1)
End If
End Sub

Private Sub GoHome_Click()
    DoCmd.OpenForm " FAMainMenu"
End Sub

Private Sub ActualDelDateSubmit_Click()
    On Error GoTo FixMe
    SQLz = "UPDATE Units INNER JOIN Bulk_Update_WorkTable ON Units.UnitID = Bulk_Update_WorkTable.UnitID SET Units.UnitActualDelDate = '" & Me.ActualDelDate & "';"
    DoCmd.RunSQL (SQLz)
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub Form_Dirty(Cancel As Integer)
   ModMade = -1
End Sub

Private Sub UnitID_DblClick(Cancel As Integer)
    Call OpenUnitsForm(Me.UnitID)
    'DoCmd.OpenForm "UnitsForm", , , "UNITID=" & Me.UnitID
    Me.UnitInServ.SetFocus
End Sub

Private Sub UnitPDStartDateSubmit_Click()
    On Error GoTo FixMe
    If IsDate(PDStartDateCBO) Then
    SQLz = "UPDATE Units INNER JOIN Bulk_Update_WorkTable ON Units.UnitID = Bulk_Update_WorkTable.UnitID SET Units.UnitPDStartDate = '" & Me.PDStartDateCBO & "';"
    DoCmd.RunSQL (SQLz)
    End If
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub InServiceSubmit_Click()
    On Error GoTo FixMe
    If Me.InservDate & "test" = "test" Then
        Call ConfrimIT("InservDate", "In Service Date")
        SQLz = "UPDATE Units INNER JOIN Bulk_Update_WorkTable " & _
               "ON Units.UnitID = Bulk_Update_WorkTable.UnitID SET Units.UnitInServ = null;"
    Else
        SQLz = "UPDATE Units INNER JOIN Bulk_Update_WorkTable " & _
               "ON Units.UnitID = Bulk_Update_WorkTable.UnitID SET Units.UnitInServ = '" & Me.InservDate & "' " & _
               "WHERE Units.UnitInServ is null;"
    End If
    DoCmd.RunSQL (SQLz)
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub


Private Sub PDRentSubmit_Click()
If Me.PDRentCurr & "test" = "test" Then Exit Sub
    On Error GoTo FixMe
    SQLz = "UPDATE Units " & _
           "SET Units.UNITPDRENT ='" & Me.PDRentCurr & "' " & _
           "WHERE GroupRef=" & Me.GroupID & ";"
    DoCmd.RunSQL (SQLz)
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub
Private Sub PDSurrenderSubmit_Click()
If Me.PDSurrenderCurr & "test" = "test" Then Exit Sub
    On Error GoTo FixMe
    SQLz = "UPDATE Units " & _
           "SET Units.PerDiemSurrender ='" & Me.PDSurrenderCurr & "' " & _
           "WHERE GroupRef=" & Me.GroupID & ";"
    DoCmd.RunSQL (SQLz)
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Public Sub ConfrimIT(ByVal FieldBoX As String, ByVal TragetX As String)
    Dim Committ As String
    If FieldBoX & "TEST" = "TEST" Then
       Committ = MsgBox("You're about set the " & TragetX & " to Null, for ALL units listed." & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Clear " & TragetX & "?")
       If Committ = vbNo Then Exit Sub
    Else
        Committ = MsgBox("You're about set the " & TragetX & vbNewLine & "to use: " & vbNewLine & _
        "" & vbNewLine & FieldBoX & vbNewLine & vbNewLine & "for ALL units listed... Do you agree?", vbQuestion + vbYesNo + vbCritical, "Change " & TragetX & "?")
    End If
End Sub
Private Sub Use_AcceptedDate_Click()
    On Error GoTo FixMe
    'If Me.AcceptedDate & "test" <> "test" Then
     Call ConfrimIT("Accepted Date", "PerDiem Start Date")
    SQLz = "UPDATE Units " & _
           "SET Units.UnitPDStartDate = UnitAcceptDate " & _
           "WHERE GroupRef=" & Me.GroupID & ";"
    DoCmd.RunSQL (SQLz)
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub Use_ActualDelDate_Click()
    On Error GoTo FixMe
     Call ConfrimIT("Delivery Date", "PerDiem Start Date")
    SQLz = "UPDATE Units " & _
           "SET Units.UnitPDStartDate = UnitActualDelDate " & _
           "WHERE GroupRef=" & Me.GroupID & ";"
    DoCmd.RunSQL (SQLz)
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub Use_InService_Click()
    On Error GoTo FixMe
     Call ConfrimIT("In Service Date", "PerDiem Start Date")
     SQLz = "UPDATE Units " & _
            "SET Units.UnitPDStartDate = UnitInServ " & _
            "WHERE GroupRef=" & Me.GroupID & ";"
    DoCmd.RunSQL (SQLz)
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub Use_VendorPaid_Click()
    On Error GoTo FixMe
     Call ConfrimIT("Vendor Paid Date", "PerDiem Start Date")
     SQLz = "UPDATE Units " & _
            "SET Units.UnitPDStartDate = UnitVendorPytDate " & _
            "WHERE GroupRef=" & Me.GroupID & ";"
    DoCmd.RunSQL (SQLz)
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub
