Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Public TimeX As Date
Private Sub Form_Load()
TimeX = Now
End Sub

Private Sub Form_Timer()
If IsNull(Me.OpenArgs) Then Me.Text0 = "a Tool Tip would show here" Else Me.Text0 = Me.OpenArgs
Me.Text2.SetFocus
If DateDiff("s", TimeX, Now) > 2 Then
Call HideControlTip
'DoCmd.close acForm, Me.Form.Name
End If
End Sub
