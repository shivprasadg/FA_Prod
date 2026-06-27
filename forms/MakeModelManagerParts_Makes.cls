Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub DeletedMe_Click()
If IsNull(Me.ChildPartMakeID) Then Exit Sub
 Dim SQLz As String
 Dim Committ, Msgx As String
 Dim Childx, ModelX As Integer
 Childx = DCount("ChildPartId", "ChildParts", "[ChildMake]= '" & Me.ChildPartMake & "'")
 ModelX = DCount("ChildPartMakeID", "LookUp_ChildPartModel", "ChildPartMakeID= " & Me.ChildPartMakeID & "")

 If ModelX > 0 Then Msgx = Msgx & ModelX & " Child Models | "
 If Childx > 0 Then Msgx = Msgx & Childx & " Child Parts | "

 If Msgx <> "" Then

 MsgBox "You Cannot Delete Make: '" & Me.ChildPartMake & "'" & vbNewLine & "It is in use by | " & Msgx, vbCritical, "Make In Use"
 Exit Sub
 End If
        Committ = MsgBox("You're about to Delete " & Me.ChildPartMake & vbNewLine & "Are you Sure?", vbQuestion + vbYesNo + vbCritical, "Delete!")
        If Committ = vbNo Then
        Exit Sub
    Else
    SQLz = "Delete * FROM LookUp_ChildPartMake WHERE ChildPartMakeID = " & Me.ChildPartMakeID & ""
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    'ClassList_Click
    'Me.Form.Refresh
        End If
End Sub

Private Sub SaveUpdate_Click()
Me.Form.Refresh
Me.Parent.MakeList.Requery
End Sub
