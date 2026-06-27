Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz As String

Private Sub ClassList_Click()
SQLz = "Select * From lookup_groupTypes WHERE [ClassID]=" & Me.ClassList.Column(0) & " ORDER BY groupType ;"
Me.Form.RecordSource = SQLz
Me.Detail.Visible = True
End Sub

Private Sub CloseMe_Click()
DoCmd.close acForm, (Me.Form.Name)
End Sub

Private Sub DeletedMe_Click()
If IsNull(Me.groupTypeID) Then Exit Sub
 Dim Committ, Typex, Msgx As String
 Dim UnitX, Childx As Integer
 UnitX = DCount("UnitId", "Units", "[UnitType]= '" & Me.groupType & "'")
 Childx = DCount("ChildPartId", "ChildParts", "[ChildType]= '" & Me.groupType & "'")
 If UnitX > 0 Then Msgx = UnitX & " Units | "
 If Childx > 0 Then Msgx = Msgx & Childx & " Child Parts |"
 If UnitX > 0 Or Childx > 0 Then

 MsgBox "You Cannot Delete '" & Me.groupType & "'" & vbNewLine & "It is in use by | " & Msgx, vbCritical, "Type In Use"
 Exit Sub
 End If
        Committ = MsgBox("You're about to Delete " & Me.groupType & vbNewLine & "Are you Sure?", vbQuestion + vbYesNo + vbCritical, "Delete!")
        If Committ = vbNo Then
        Exit Sub
    Else
    SQLz = "Delete * FROM lookup_groupTypes WHERE GroupTypeID = " & Me.groupTypeID & ""
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    ClassList_Click
    'Me.Form.Refresh
        End If
End Sub

Private Sub groupType_AfterUpdate()
If Me.Form.NewRecord Then
Me.GroupClass = Me.ClassList.Column(1)
Me.ClassID = Me.ClassList.Column(0)
End If
End Sub

Private Sub SaveUpdate_Click()
Me.Form.Refresh
End Sub
