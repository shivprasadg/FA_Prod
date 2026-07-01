Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz As String
Dim SchIDx As Long

Private Sub CloseMe_Click()
    DoCmd.OpenForm "FAMainMenu"
'    If ModMade = -1 Then Call CheckATFChanges(Me.groupID, "Rent Change")
    DoCmd.close acForm, Me.Form.Name
End Sub


Private Sub Form_Current()
    On Error GoTo whoops
    Me.txtCurrent = Me.GroupID
whoops:
    Exit Sub
End Sub



Private Sub Form_Load()
    SchIDx = Nz(Me.OpenArgs, 0)
    SQLz = "SELECT Clients.ClientGroupID, Clients.clientID, Schedule.SchID, SchGrp.SGrpID, SchGrp.UnitsInGroup, Schedule.SchNo, SchGrp.sGrpNo"
    SQLz = SQLz & " , SchGrp.PerDiemStartsOn, SchGrp.BLCD, SchGrp.SGRPNoPD, Clients.ClientShNm, SchGrp.SgrpCommInd"
    SQLz = SQLz & " FROM ((Clients "
    SQLz = SQLz & " INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client) "
    SQLz = SQLz & " INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef) "
    SQLz = SQLz & " INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID"
    SQLz = SQLz & " WHERE (((Schedule.SchID)=" & SchIDx & ")) ORDER BY sGrpNo;"
    Me.Form.RecordSource = SQLz
End Sub

Private Sub GoHome_Click()
    DoCmd.OpenForm " FAMainMenu"
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


Private Sub PerDiemStartsOnSubmitt_Click()
    Call ConfrimIT("Perdiem Starts On", "PerDiem Start Type")
    SQLz = "UPDATE SchGrp " & _
        "SET UPerDiemStartsOn = '" & PerDiemStartsOnCBO & "' " & _
        "WHERE SGrpSchID=" & Me.SchID & ";"
    DoCmd.RunSQL (SQLz)
FixMe:
    Me.Form.Refresh
End Sub

Private Sub SgrpCommIndSubmitt_Click()
    Call ConfrimIT("BLCD Starts On", "BLCD Start Type")
    SQLz = "UPDATE SchGrp " & _
        "SET SgrpCommInd = '" & SgrpCommIndCBO & "' " & _
        "WHERE SGrpSchID=" & Me.SchID & ";"
    DoCmd.RunSQL (SQLz)
FixMe:
    Me.Form.Refresh
End Sub

Private Sub SGRPNoPDSubmitt_Click()
    Call ConfrimIT("NO Perdiem", "NO Perdiem Yes/No")
    SQLz = "UPDATE SchGrp " & _
        "SET SGRPNoPD = " & SGRPNoPDCBO & " " & _
        "WHERE SGrpSchID=" & Me.SchID & ";"
    DoCmd.RunSQL (SQLz)
FixMe:
    Me.Form.Refresh
End Sub
