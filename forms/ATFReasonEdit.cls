Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
Me.info.SetFocus
End Sub

Private Sub SaveChanges_Click()
    If Len(Me.ReasonDescription) < 10 Then
        Me.Form.Undo
        MsgBox "You must have a valid Reason", vbCritical, AtlaasVersion
    End If
    DoCmd.RunCommand acCmdSaveRecord
    If IsOpen("GroupForm") Then
        Forms!GroupForm.ReasonNote = Me.ReasonDescription
    End If
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub Undo_Click()
Me.Form.Undo
End Sub
