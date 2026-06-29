Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim ii         As Integer
Dim Msgx, ResposeX As String

Private Sub AddNew_Click()
If DCount("ClientID", "SUFConditions", "ClientID=" & Nz(Me.ClientID, 0) & "") > 0 Then
        MsgBox "There already are SUF Conditions attached to this client", vbCritical, "There are SUF Conditions "
Exit Sub
End If
    Dim SQLz   As String
    Msgx = "<Font Size=4><b>Adding SUF Conditions</b></font><br /><br/>"
    Msgx = Msgx & HSpaces(3) & "You're about to add SUF Conditons<br/>"
    Msgx = Msgx & HSpaces(3) & "For Client:<b> " & Me.Parent.ClientShNm & "</b><br/>"
    Msgx = Msgx & HSpaces(3) & "This will be a SUF Condition 'template'<br/>"
    Msgx = Msgx & HSpaces(3) & "You may need to modify it for this Client <br/><br/><b>"
    Msgx = Msgx & HSpaces(3) & "Do You Want To Proceed?<br/>"
    ResposeX = Dialog.RichBox(Msgx, vbCustom + vbExclamation, AtlaasVersion, , , 0, 1, 0, 0, "Cancel", "Proceed")
    If ResposeX = 101 Then Exit Sub
    Me.LastUpdated.SetFocus
    SQLz = SQLz & "INSERT INTO [SUFConditions]"
    SQLz = SQLz & "(Condition1"
    SQLz = SQLz & ",Condition2"
    SQLz = SQLz & ",Condition3"
    SQLz = SQLz & ",Condition4"
    SQLz = SQLz & ",Condition5"
    SQLz = SQLz & ",Condition6"
    SQLz = SQLz & ",Condition7"
    SQLz = SQLz & ",Condition8"
    SQLz = SQLz & ",Condition9"
    SQLz = SQLz & ",Condition10"
    SQLz = SQLz & ",ClientID"
    SQLz = SQLz & ",VechicleType"
    SQLz = SQLz & ",UpdateBy"
    SQLz = SQLz & ",LastUpdated) "
    SQLz = SQLz & "Select Condition1"
    SQLz = SQLz & ",Condition2"
    SQLz = SQLz & ",Condition3"
    SQLz = SQLz & ",Condition4"
    SQLz = SQLz & ",Condition5"
    SQLz = SQLz & ",Condition6"
    SQLz = SQLz & ",Condition7"
    SQLz = SQLz & ",Condition8"
    SQLz = SQLz & ",Condition9"
    SQLz = SQLz & ",Condition10"
    SQLz = SQLz & "," & Me.Parent.ClientID & ""
    SQLz = SQLz & ",VechicleType"
    SQLz = SQLz & ",'" & FAUserName & "'"
    SQLz = SQLz & ",'" & Date & "'"
    SQLz = SQLz & "FROM SUFConditions WHERE SUFCondID=1"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Me.Form.Requery
    Me.LastUpdated.SetFocus
    Call ExpandContract(0)
End Sub

Private Sub DeleteConditons_Click()
    If IsNull(Me.SUFCondID) Then Exit Sub
    Msgx = "<Font Size=4><b>Deleting SUF Conditons</b></font><br /><br/>"
    Msgx = Msgx & HSpaces(3) & "You're about to<b> Delete</b> SUF Conditons<br/>"
    Msgx = Msgx & HSpaces(3) & "For Client:<b> " & Me.Parent.ClientShNm & "</b><br/>"
    Msgx = Msgx & HSpaces(3) & "This<b> Cannot</b> be undone<br/><br/>"
    Msgx = Msgx & HSpaces(3) & "<b>Do You Want To Proceed?<br/>"
    ResposeX = Dialog.RichBox(Msgx, vbCustom + vbExclamation, AtlaasVersion, , , 0, 1, 0, 0, "Cancel", "Proceed")
    If ResposeX = 101 Then Exit Sub
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete * FROM SUFConditions WHERE SUFCondID=" & Me.SUFCondID & "")
    DoCmd.SetWarnings True
    Me.Form.Requery
    Me.LastUpdated.SetFocus
    Call ExpandContract(0)
End Sub

Public Sub SetSuf()
 Dim Lockx  As Boolean
 Me.Form.RecordSource = "Select * From SUFConditions Where ClientId=" & Me.Parent.ClientID & ";"
    If FAUserAccessLevel < 15 Then Lockx = True Else Lockx = False
    For ii = 1 To 10
        Me("Condition" & ii).Locked = Lockx
    Next ii
    Call ExpandContract(0)
End Sub

Private Sub SaveMe_Click()
    Me.Form.Refresh
End Sub

Private Sub Form_Close()
    On Error GoTo Form_Close_Error
    On Error GoTo 0
    Exit Sub
Form_Close_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Close of VBA Document " & Me.Form.Name
End Sub

Private Function ExpandContract(Boxi As Integer)
    For ii = 1 To 10
        If Boxi <> ii Then
            Me("Condition" & ii).Height = (0.16 * 1440)
            ' Me("Condition" & ii).Visible = False
        End If
        If Boxi = 0 Then
            Me("Condition" & ii).Height = (1 * 1440)
            ' Me("Condition" & ii).Visible = True
        End If
    Next ii
    For ii = 1 To 10
        If Boxi = ii And Boxi > 0 Then
            Me("Condition" & ii).Height = (4.4 * 1440)
            Me("Condition" & ii).Visible = True
        End If
    Next ii
End Function

Private Sub SUFReportView_Click()
    Dim UnitIDx As Long
    Dim Queryx As String
    Dim ReportX As String
    Dim SQLy   As String
    On Error Resume Next
    UnitIDx = Nz(DLookup("UnitId", "Units", "UnitStatus = 'A' AND UnitType ='" & Me.VechicleType & "' AND ClientID=" & Me.ClientID & ""), 0)
    On Error GoTo 0
    If UnitIDx = 0 Then
        MsgBox "There are no active units for this client", vbCritical, "No Active Units"
        Exit Sub
    End If
    If IsNull(UnitIDx) Then
        MsgBox "You must Select Units before you can request an SUF", vbCritical, "No Units Selected"
        Exit Sub
    End If
    Queryx = "vw_SUFReportQuery"
    ReportX = "SUFwConditions"
    If Me.Parent.ClientGroupID = 27 Then ReportX = "SUFReport_Maines"
    SQLy = "SELECT * "
    SQLy = SQLy & "FROM " & Queryx & " "
    SQLy = SQLy & "WHERE UnitID=" & UnitIDx & " "
    SQLy = SQLy & "ORDER BY UnitVin;"
    If ReportX = "" Then Exit Sub
    DoCmd.close acReport, ReportX
    DoCmd.OpenReport ReportX, acViewReport, , , , SQLy
End Sub
Private Sub Condition1_LostFocus()
    Call ExpandContract(0)
End Sub
Private Sub Condition1_GotFocus()
    Call ExpandContract(1)
End Sub
Private Sub Condition2_LostFocus()
    Call ExpandContract(0)
End Sub
Private Sub Condition2_GotFocus()
    Call ExpandContract(2)
End Sub
Private Sub Condition3_LostFocus()
    Call ExpandContract(0)
End Sub
Private Sub Condition3_GotFocus()
    Call ExpandContract(3)
End Sub
Private Sub Condition4_LostFocus()
    Call ExpandContract(0)
End Sub
Private Sub Condition4_GotFocus()
    Call ExpandContract(4)
End Sub
Private Sub Condition5_LostFocus()
    Call ExpandContract(0)
End Sub
Private Sub Condition5_GotFocus()
    Call ExpandContract(5)
End Sub
Private Sub Condition6_GotFocus()
    Call ExpandContract(6)
End Sub
Private Sub Condition6_LostFocus()
    Call ExpandContract(0)
End Sub
Private Sub Condition7_GotFocus()
    Call ExpandContract(7)
End Sub
Private Sub Condition7_LostFocus()
    Call ExpandContract(0)
End Sub
Private Sub Condition8_GotFocus()
    Call ExpandContract(8)
End Sub
Private Sub Condition8_LostFocus()
    Call ExpandContract(0)
End Sub
Private Sub Condition9_GotFocus()
    Call ExpandContract(9)
End Sub
Private Sub Condition9_LostFocus()
    Call ExpandContract(0)
End Sub
Private Sub Condition10_GotFocus()
    Call ExpandContract(10)
End Sub
Private Sub Condition10_LostFocus()
    Call ExpandContract(0)
End Sub
Private Sub ExpandAll_Click()
    Call ExpandContract(0)
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)

    Select Case DataErr
    Case 3146    'Required Field ODBC Error 'Case 3146 To 3621 'ODBC Error
        MsgBox "You must have data in all Required Fields before moving off the record or closing the form.", vbCritical, "Required Field Error #" + str(DataErr) + " - All Required Fields Must Contain Data!"
        Response = acDataErrContinue
    Case 3146 To 3621
        'MsgBox "Unknown ODBC error occurred", vbCritical, "ODBC Error #" & Str(DataErr)
        'Response = acDataErrContinue
    Case Else
        MsgBox AccessError(DataErr)
        Response = acDataErrContinue
    End Select
End Sub
