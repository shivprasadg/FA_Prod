Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Label52_Click()
If Me.SchID <> 0 Then Call Schedule_Click
End Sub

Private Sub Schedule_Click()
If Me.SchID = 0 Then Exit Sub
MessageUser "Open", "Opening Schedule Form", "just a moment.."
DoCmd.OpenForm "Schedule_Form", , , "SchID =" & Me.SchID
MessageUser "Close"
End Sub
