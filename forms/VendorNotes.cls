Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub AddNoteBtn_Click()
 DoCmd.GotoRecord , , acNewRec
 Me.vnVendorID = Me.Parent.VendorID
 Me.vnFAUser = FAUserInit
 Me.vnDate = Date
 Me.vnMemo.SetFocus
End Sub


Private Sub Form_Current()
    Me.NotesList.Requery
End Sub

Private Sub NotesList_Click()
On Error GoTo Fxme
Me.Form.FilterOn = False
 Set RST = Me.RecordsetClone
       RST.FindFirst "VendorNoteID = " & Me.NotesList.Column(0) & ""
       Me.Bookmark = RST.Bookmark
    RST.close
    Set RST = Nothing
Exit Sub
Fxme:
MsgBox "No Matching Record", vbCritical, AtlaasVersion
RST.close
Set RST = Nothing
End Sub

Private Sub SaveMe_Click()
    If Me.Dirty Then
        Me.Dirty = False
        Me.NotesList.Requery
    End If
End Sub
