Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public msg1 As String
Public ResposeX As VbMsgBoxResultEx

Private Sub BillType_Click()
    Me.BillTypeID = Me.BillType.Column(1)
End Sub

Private Sub CloseMe_Click()
    Call ValidateBill
    If msg1 <> "" Then Exit Sub
    If Me.Form.Dirty Then Me.Form.Dirty = False
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub Form_Close()
    ' Make sure that there is a BillTypeID assigned to the Bill Type.
    '  Dim SQLz As String
    '  SQLz = "UPDATE B SET B.BillTypeID = T.BillTypeID"
    '  SQLz = SQLz & " FROM Bills AS B"
    '  SQLz = SQLz & " INNER JOIN Lookup_BillType AS T on B.BillType = T.BillType"
    '  SQLz = SQLz & " WHERE B.BillTypeID Is Null"
On Error GoTo FixMe
    '  DoCmd.SetWarnings False
    '  DoCmd.RunSQL (SQLz)
    '  DoCmd.SetWarnings True
    '  Exit Sub
FixMe:
    DoCmd.SetWarnings True
    Me.Form.Undo
    MsgBox Err.Description, vbCritical, "(Error!)" & AtlaasVersion
End Sub
Public Sub ValidateBill()
    msg1 = ""
    If Me.BillType & "test" = "test" Or Me.BillTypeID & "test" = "test" Then msg1 = msg1 & "You did not Select a <b>Bill Type:</b><br/>"
    If Me.Billdt & "test" = "test" Then msg1 = msg1 & "You did not select a Bill <b>Date:</b><br/>"
    If Me.BillNo & "test" = "test" Then msg1 = msg1 & "You did not set a Bill <b>Type:</b><br/>"
    If Me.Billfrom & "test" = "test" Then msg1 = msg1 & "You did not set a Bill <b>From:</b><br/>"
    If Me.BillDueDt & "test" = "test" Then msg1 = msg1 & "You did not set a Bill <b>Due:</b><br/>"
    If msg1 <> "" Then
        msg1 = "<b><Font Size=3>There is an Error with this Bill</font></b><br/>" & msg1
        ResposeX = Dialog.RichBox(msg1, vbCustom + vbExclamation, AtlaasVersion, , , 0, 1, 0, 0, "OK")
    End If
End Sub
Private Sub SaveMe_Click()
    Call ValidateBill
    If msg1 <> "" Then Exit Sub
    If Me.Form.Dirty Then Me.Form.Dirty = False
End Sub
