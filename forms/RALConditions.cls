Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim ii         As Integer
Dim Msgx, ResposeX As String

Private Sub AddNew_Click()
If DCount("BankID", "RALParagraphs", "BankID=" & Nz(Me.BankID, 0) & "") > 0 Then
        MsgBox "There already are RAL Paragraphs attached to this Bank", vbCritical, "There are RAL Paragraphs "
Exit Sub
End If
    Dim SQLz   As String
    Msgx = "<Font Size=4><b>Adding RAL Paragraphs</b></font><br /><br/>"
    Msgx = Msgx & HSpaces(3) & "You're about to add RAL Conditons<br/>"
    Msgx = Msgx & HSpaces(3) & "For Client:<b> " & Me.Parent.ClientShNm & "</b><br/>"
    Msgx = Msgx & HSpaces(3) & "This will be a RAL Paragraph 'template'<br/>"
    Msgx = Msgx & HSpaces(3) & "You may need to modify it for this Bank <br/><br/><b>"
    Msgx = Msgx & HSpaces(3) & "Do You Want To Proceed?<br/>"
    ResposeX = Dialog.RichBox(Msgx, vbCustom + vbExclamation, AtlaasVersion, , , 0, 1, 0, 0, "Cancel", "Proceed")
    If ResposeX = 101 Then Exit Sub
    Me.LastUpdated.SetFocus
    SQLz = SQLz & "INSERT INTO [RALParagraphs]"
    SQLz = SQLz & "(Paragraph1"
    SQLz = SQLz & ",Paragraph2"
    SQLz = SQLz & ",Paragraph3"
    SQLz = SQLz & ",Paragraph4"
    SQLz = SQLz & ",Paragraph5"
    SQLz = SQLz & ",Paragraph6"
    SQLz = SQLz & ",Paragraph7"
    SQLz = SQLz & ",Paragraph8"
    SQLz = SQLz & ",Paragraph9"
    SQLz = SQLz & ",Paragraph10"
    SQLz = SQLz & ",BankID"
    SQLz = SQLz & ",VechicleType"
    SQLz = SQLz & ",UpdateBy"
    SQLz = SQLz & ",LastUpdated) "
    SQLz = SQLz & "Select Paragraph1"
    SQLz = SQLz & ",Paragraph2"
    SQLz = SQLz & ",Paragraph3"
    SQLz = SQLz & ",Paragraph4"
    SQLz = SQLz & ",Paragraph5"
    SQLz = SQLz & ",Paragraph6"
    SQLz = SQLz & ",Paragraph7"
    SQLz = SQLz & ",Paragraph8"
    SQLz = SQLz & ",Paragraph9"
    SQLz = SQLz & ",Paragraph10"
    SQLz = SQLz & "," & Me.Parent.BankID & ""
    SQLz = SQLz & ",VechicleType"
    SQLz = SQLz & ",'" & FaUserName & "'"
    SQLz = SQLz & ",'" & Date & "'"
    SQLz = SQLz & "FROM RALConditions WHERE RALCondID=1"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Me.Form.Requery
    Me.LastUpdated.SetFocus
    Call ExpandContract(0)
End Sub

Private Sub DeleteConditons_Click()
    If IsNull(Me.RALCondID) Then Exit Sub
    Msgx = "<Font Size=4><b>Deleting RAL Conditons</b></font><br /><br/>"
    Msgx = Msgx & HSpaces(3) & "You're about to<b> Delete</b> RAL Conditons<br/>"
    Msgx = Msgx & HSpaces(3) & "For Client:<b> " & Me.Parent.ClientShNm & "</b><br/>"
    Msgx = Msgx & HSpaces(3) & "This<b> Cannot</b> be undone<br/><br/>"
    Msgx = Msgx & HSpaces(3) & "<b>Do You Want To Proceed?<br/>"
    ResposeX = Dialog.RichBox(Msgx, vbCustom + vbExclamation, AtlaasVersion, , , 0, 1, 0, 0, "Cancel", "Proceed")
    If ResposeX = 101 Then Exit Sub
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete * FROM RALParagraphs WHERE RALCondID=" & Me.RALCondID & "")
    DoCmd.SetWarnings True
    Me.Form.Requery
    Me.LastUpdated.SetFocus
    Call ExpandContract(0)
End Sub

Public Sub SetRAL()
 Dim Lockx  As Boolean
 Me.Form.RecordSource = "Select * From RALParagraphs Where BankID=" & Me.Parent.BankID & ";"
    If FAUserAccessLevel < 15 Then Lockx = True Else Lockx = False
    For ii = 1 To 10
        Me("Paragraph" & ii).Locked = Lockx
    Next ii
    Call ExpandContract(0)
End Sub

Private Sub Form_Open(Cancel As Integer)

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
            Me("Paragraph" & ii).Height = (0.16 * 1440)
            ' Me("Paragraph" & ii).Visible = False
        End If
        If Boxi = 0 Then
            Me("Paragraph" & ii).Height = (1 * 1440)
            ' Me("Paragraph" & ii).Visible = True
        End If
    Next ii
    For ii = 1 To 10
        If Boxi = ii And Boxi > 0 Then
            Me("Paragraph" & ii).Height = (4.4 * 1440)
            Me("Paragraph" & ii).Visible = True
        End If
    Next ii
End Function

Private Sub RALReportView_Click()
    Dim UnitIDx As Long
    Dim Queryx As String
    Dim ReportX As String
    Dim SQLy   As String
    On Error Resume Next
    UnitIDx = Nz(DLookup("UnitId", "Units", "UnitStatus = 'A' AND UnitType ='" & Me.VechicleType & "' AND BankID=" & Me.BankID & ""), 0)
    On Error GoTo 0
    If UnitIDx = 0 Then
        MsgBox "There are no active units for this Bank", vbCritical, "No Active Units"
        Exit Sub
    End If
    If IsNull(UnitIDx) Then
        MsgBox "You must Select Units before you can request an RAL", vbCritical, "No Units Selected"
        Exit Sub
    End If
    Queryx = "vw_RALReportQuery"
    ReportX = "RALwParagraphs"
    If Me.Parent.ClientGroupID = 27 Then ReportX = "RALReport_Maines"
    SQLy = "SELECT * "
    SQLy = SQLy & "FROM " & Queryx & " "
    SQLy = SQLy & "WHERE UnitID=" & UnitIDx & " "
    SQLy = SQLy & "ORDER BY UnitVin;"
    If ReportX = "" Then Exit Sub
    DoCmd.close acReport, ReportX
    DoCmd.OpenReport ReportX, acViewReport, , , , SQLy
End Sub
Private Sub Paragraph1_LostFocus()
    Call ExpandContract(0)
End Sub
Private Sub Paragraph1_GotFocus()
    Call ExpandContract(1)
End Sub
Private Sub Paragraph2_LostFocus()
    Call ExpandContract(0)
End Sub
Private Sub Paragraph2_GotFocus()
    Call ExpandContract(2)
End Sub
Private Sub Paragraph3_LostFocus()
    Call ExpandContract(0)
End Sub
Private Sub Paragraph3_GotFocus()
    Call ExpandContract(3)
End Sub
Private Sub Paragraph4_LostFocus()
    Call ExpandContract(0)
End Sub
Private Sub Paragraph4_GotFocus()
    Call ExpandContract(4)
End Sub
Private Sub Paragraph5_LostFocus()
    Call ExpandContract(0)
End Sub
Private Sub Paragraph5_GotFocus()
    Call ExpandContract(5)
End Sub
Private Sub Paragraph6_GotFocus()
    Call ExpandContract(6)
End Sub
Private Sub Paragraph6_LostFocus()
    Call ExpandContract(0)
End Sub
Private Sub Paragraph7_GotFocus()
    Call ExpandContract(7)
End Sub
Private Sub Paragraph7_LostFocus()
    Call ExpandContract(0)
End Sub
Private Sub Paragraph8_GotFocus()
    Call ExpandContract(8)
End Sub
Private Sub Paragraph8_LostFocus()
    Call ExpandContract(0)
End Sub
Private Sub Paragraph9_GotFocus()
    Call ExpandContract(9)
End Sub
Private Sub Paragraph9_LostFocus()
    Call ExpandContract(0)
End Sub
Private Sub Paragraph10_GotFocus()
    Call ExpandContract(10)
End Sub
Private Sub Paragraph10_LostFocus()
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
