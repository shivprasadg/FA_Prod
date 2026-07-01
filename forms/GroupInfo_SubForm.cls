Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub GroupNum_Click()
MessageUser "Open", "Opening Groups Form", "just a moment.."
DoCmd.OpenForm "GroupForm", , , "SGrpID =" & Me.GroupID
MessageUser "Close"
End Sub

Private Sub Label52_Click()
If Me.GroupID <> 0 Then Call GroupNum_Click
End Sub
