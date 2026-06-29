Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim RST        As DAO.Recordset

Private Sub CleanUpNotes_Click()
Call MessageUser("Open", "Cleaning Up...", "Just a moment")
Call CleanRevComments
Me.Form.Requery
Call MessageUser("Close")
End Sub

Private Sub RevID_AfterUpdate()
If DCount("RevID", "ATF_ChangeReasons", "SchId=" & Me.SchID & " And RevID=" & Me.RevID & "") > 0 Then
DoCmd.Beep
Me.RevID = NextRevID(Me.SchID)
Me.NewValue = "Note " & Me.RevID
DoCmd.Beep
End If
End Sub
Private Sub AddRevision_Click()
   ' On Error GoTo Fixme
    If AskToDoTask("You're about to add a ATF Revision Note.") = False Then Exit Sub
    'If Me(FieldBoX).Value & "TEST" = "TEST" Then
        DoCmd.GotoRecord , , acNewRec
        Me.RevID = Nz(NextRevID(Me.Parent.SchID), 0)
        Me.Action = "ATF Note " & Me.RevID
        Me.ChangeDate = Now()
        Me.SysUserCbo = FAUserName
        Me.NewValue = "Note " & Me.RevID
        Me.FieldGroup = "Schedule Notes"
        Me.FieldName = "Rev-Notes"
        Me.SchNo = Me.Parent.SchNo
        Me.Reason = "|...type notes here...|"
        Me.Reason.SetFocus
    Exit Sub
FixMe:
    MsgBox "There was an error"
End Sub

Private Sub DeleteRev_Click()
    On Error GoTo FixMe
    If AskToDoTask("You're about to delete a ATF/Schedule Revision") = True Then
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete From ATF_ChangeReasons WHERE Indx=" & Me.INDX)
    DoCmd.SetWarnings True
    Me.Form.Requery
    Me.RevList.Requery
    End If
FixMe:
End Sub

Private Sub Form_Dirty(Cancel As Integer)
    Me.Undo
End Sub

Private Sub Form_Load()
    If FAUserAccess = 99 Or FAUserAccess = 17 Then
        Me.DeleteRev.Visible = True
        Me.CleanUpNotes.Visible = True
        Me.RevID.Locked = False
    Else
        Me.DeleteRev.Visible = False
        Me.CleanUpNotes.Visible = False
        Me.RevID.Locked = True
    End If
End Sub

Private Sub Reason_Change()
If Me.Reason.Text Like "*|...type notes here...|*" Then
Me.Reason.Text = SplitString(Me.Reason.Text, "here...|", 1)
End If
Me.Reason.SetFocus
End Sub
Public Sub RevList_Click()
    Me.Form.FilterOn = False
    Set RST = Me.RecordsetClone
    RST.FindFirst "INDX = " & Me.RevList.Column(0) & ""
    Me.Bookmark = RST.Bookmark
    Set RST = Nothing
End Sub

Private Sub SaveUpdate_Click()
If Me.Reason Like "*|...type notes here...|*" Or Len(Nz(Me.Reason, "")) < 10 Then
MsgBox "You must enter some sort of notes to proceed", vbCritical, AtlaasVersion
Me.Reason = "|...type notes here...|"
Me.Reason.SetFocus
Exit Sub
End If
    Me.Form.Refresh
End Sub
