Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub CloseMe_Click()
On Error GoTo FixMe
DoCmd.close acForm, Me.Form.Name
Exit Sub
FixMe:
MsgBox "A Schedule was NOT Added", vbCritical, "Un-Done"
Me.Form.Undo
DoCmd.close acForm, Me.Form.Name
End Sub
