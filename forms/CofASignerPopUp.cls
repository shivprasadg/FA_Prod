Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz As String
Private Sub CloseMe_Click()
    On Error GoTo FixMe
    DoCmd.close acForm, Me.Form.Name
    Exit Sub
FixMe:
    MsgBox "A C of A was NOT completed", vbCritical, "Un-Done"
    Me.Form.Undo
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub CofaSignCbo_Click()
    On Error GoTo FixMe
    Dim Committ As String

    Committ = MsgBox("You're about to set " & Me.CofaSignCbo.Column(1) & " as the" & vbNewLine & "CofA Signer for all (" & Me.UnitsInGroup & ") units in this group" & vbNewLine & "Are you sure?", vbQuestion + vbYesNo + vbCritical, "Set CofA Signer")
    If Committ = vbNo Then
        Exit Sub
    Else
        SQLz = "UPDATE Units SET Units.CofASigner = " & Me.CofaSignCbo.Column(0) & " WHERE Units.GroupREF=" & Me.groupID & ";"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
    End If
    Exit Sub
FixMe:
    MsgBox "There was an error"
End Sub


Private Sub DeliveryAsAccepted_Click()
    If Me.DeliveryAsAccepted = True Then
        On Error GoTo FixMe
        Dim Committ As String
        Committ = MsgBox("You're about to set the Delivery Date as the Accepted Date" _
                     & vbNewLine & "for this CofA all (" & Me.UnitsInGroup & ") units in this group" _
                     & vbNewLine & "Are you sure?", vbQuestion + vbYesNo + vbCritical, "Set Delivery Date as Accepted")
        If Committ = vbNo Then
            Me.DeliveryAsAccepted = False
            Exit Sub
        Else
            SQLz = "UPDATE Units SET UnitAcceptDate = UnitActualDelDate " & _
                   "WHERE NOT UnitActualDelDate IS Null " & _
                   "AND UnitAcceptDate Is Null " & _
                   "AND Units.GroupREF=" & Me.groupID & ";"

    '       DoCmd.SetWarnings False
            DoCmd.RunSQL (SQLz)
            DoCmd.SetWarnings True
        End If
        Exit Sub
    End If
FixMe:
End Sub

Private Sub Form_Load()
'Dim RST    As DAO.Recordset
SQLz = "SELECT S.ClientGroupID, S.clientID, S.SchID, S.GroupID, S.AssetsInGroup, S.Schedule, S.UnitGroup, U.CofASigner, C.contactFullName AS SignerName "
SQLz = SQLz & "FROM (vw_SixKeys AS S INNER JOIN Units AS U ON S.AssetID = U.UnitId) Left JOIN Contacts AS C ON U.CofASigner = C.contactID "
SQLz = SQLz & "WHERE  GroupID=" & Me.OpenArgs & " "
Me.Form.RecordSource = SQLz
'SQLz = "SELECT contactID, contactFullName As Name, contactTitle as Title, contactInc, contactCofA "
'SQLz = SQLz & "FROM Contacts "
'SQLz = SQLz & "WHERE clientRef=" & Me.clientID & " "
'SQLz = SQLz & "AND ContactObsolete=No;"
Me.CofaSignCbo.Requery
'RowSource = SQLz
Me.Schedule_CoLessee.Requery
Me.CofaSignCbo.Requery
Me.AskForDate = 0
Me.HideAcceptedDate = 0
Me.HideSignatureDate = 0

End Sub

Private Sub MakeCofA_Click()
    Dim ReportX As String
    Dim Response As String

    ReportX = "CofA_Mass"
    If Me.ClientGroupID = 187 Then
        SQLz = "UPDATE Units SET UnitAcceptDate = UnitActualDelDate " & _
            "WHERE NOT UnitActualDelDate Is Null " & _
            "AND UnitActualDelDate Is Null " & _
            "AND Units.GroupREF=" & Me.groupID & ";"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
    End If
    If Me.AskForDate = 0 Then
    If Me.ClientID = 3318 Then ReportX = "CofA_Mass_AirProd"
    If Me.ClientID = 5442 Then ReportX = "CofA_Mass_AirProd"
    If Me.ClientID = 3319 Then ReportX = "CofA_Mass_AirGas"
    If Me.ClientID = 4412 Then ReportX = "CofA_Mass_Dean"
   '' If Me.clientID = 4046 Then ReportX = "CofA_Mass_Dyno"
    If Me.ClientID = 5216 Then ReportX = "CofA_Mass_Linde"
    If Me.ClientID = 5338 Then ReportX = "CofA_Mass_SPD"
    If Me.ClientID = 3674 Then ReportX = "CofA_Mass_Sygma"
    If Me.ClientID = 3686 Then ReportX = "CofA_Mass_UGI"
    If Me.ClientID = 3687 Then ReportX = "CofA_Mass_Unisource"
    End If
    If Me.AskForDate = -1 And ReportX = "CofA_Mass" Then ReportX = "CofA_Mass_wDate"
    ' Response = MsgBox("Do you want to insert a 'CofA' date?", vbQuestion + vbYesNo + vbCritical, "Will Print on Form")
    '    If Response = vbYes Then ReportX = "CofA_Mass_wDate"
    'End If
MakeReport:
    DoCmd.OpenReport ReportX, acViewPreview
ExitMe:
    DoCmd.close acForm, "CofASignerPopUp"
End Sub

Private Sub MakeCofASchedule_Click()
    Dim ReportX As String
    Dim Response As String
   ReportX = "CofA_Mass_Sch"
   If Me.AskForDate = -1 Then ReportX = "CofA_Mass_Sch_wDate"
   DoCmd.OpenReport ReportX, acViewPreview
ExitMe:
    DoCmd.close acForm, "CofASignerPopUp"
End Sub
