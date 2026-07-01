Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub CloseMe_Click()
On Error GoTo FixMe
If Me.TitleState & "test" = "test" Then
MsgBox "You must select a state", vbCritical, "No State, No Save"

Exit Sub
End If
FixMe:
DoCmd.close acForm, Me.Form.Name

End Sub

Private Sub LogTitleIn_Click()
Dim SQLz As String
SQLz = "INSERT INTO UnitTitleLog ( Unitref, TitleNumber, TitleSt, Reason, UnitTitleLogDate, EntryTypeId, UserID ) " & _
       "SELECT " & Me.UnitID & " as unitid, '" & Me.UnitTitleNumber & "' AS Tnum, '" & Me.TitleState & "' AS State, '" & Me.EntryReasonId.Column(0) & "' AS Reason, '" & Date & "', 1, ,'" & FaUserName & "';"
DoCmd.SetWarnings False
DoCmd.RunSQL (SQLz)
DoCmd.SetWarnings True
MsgBox "Title #" & Me.UnitTitleNumber & "has been checked in" & vbNewLine & " The Form will Close", , "checked In"
DoCmd.close acForm, Me.Form.Name

End Sub
