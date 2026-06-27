Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim RST As DAO.Recordset
Private Sub AddRevision_Click()
    On Error GoTo FixMe
    Dim NextNum, Committ As String
    NextNum = Nz(DMax("AmendmentNo", "ScheduleAmends", "ScheduleID=" & Me.Parent.SchID & ""), 1)
    If NextNum > 0 Then NextNum = NextNum + 1
    Committ = MsgBox("You're about to create" & vbNewLine & "Schedule Amendment" & vbNewLine & "                     Number: " & NextNum & vbNewLine & "Are you sure?", vbQuestion + vbYesNo + vbCritical, "New Amendment Comfrim")
    If Committ = vbNo Then
        Exit Sub
    Else
        DoCmd.GotoRecord , , acNewRec
        Me.AmendmentNO = NextNum
        Me.AmendmentDate = Date
        Me.Form.Refresh
        Me.AmendmentDesc.SetFocus
    End If
    Exit Sub
FixMe:
    MsgBox "There was an error"
End Sub

Private Sub DeleteSelected_Click()
If Me.RevList & "test" = "test" Then Exit Sub
Dim Committ, SQLz As String
 Committ = MsgBox("You're about to delete" & vbNewLine & "Schedule Amendment" & vbNewLine & "                     Number: " & Me.RevList.Column(2) & vbNewLine & "Are you sure?", vbQuestion + vbYesNo + vbCritical, "Comfrim Delete")
    If Committ = vbNo Then
        Exit Sub
    Else
SQLz = "Delete From ScheduleAmends Where ID = " & Me.RevList.Column(0) & ";"
DoCmd.SetWarnings False
DoCmd.RunSQL (SQLz)
DoCmd.SetWarnings True
Me.RevList.Requery
    End If
End Sub

Private Sub Form_Current()
    Me.RevList.Requery
End Sub

Private Sub RevList_Click()
    Set RST = Me.RecordsetClone
    RST.FindFirst "ID = " & Me.RevList.Column(0) & ""
    Me.Bookmark = RST.Bookmark
    RST.close
    Set RST = Nothing
End Sub

Private Sub SaveUpdate_Click()
    DoCmd.RunCommand acCmdSaveRecord
End Sub
