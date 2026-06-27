Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz       As String

Private Sub CloseMe_Click()
    DoCmd.close acForm, (Me.Form.Name)
End Sub

Private Sub DeletedMe_Click()
    If Me.Form.NewRecord = True Then Exit Sub
    Dim Committ, Typex, Msgx As String
    Dim UnitX, Childx As Integer
    If Nz(Me.Leased, 0) > 0 Or Nz(Me.Tracked, 0) > 0 Then
        MsgBox "You Cannot Delete '" & Me.UnitModel & "'" & vbNewLine & "It is in use", vbCritical, "Type In Use"
        Exit Sub
    End If
    Committ = MsgBox("You're about to Delete " & Me.UnitModel & vbNewLine & "Are you Sure?", vbQuestion + vbYesNo + vbCritical, "Delete!")
    If Committ = vbNo Then
        Exit Sub
    Else
        SQLz = "Delete * FROM lookup_UnitModels WHERE INDX = " & Me.INDX & ""
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
        Me.Form.Requery
    Me.Parent.ModelList.Requery
    Me.Parent.NoMatchModelsList.Requery
    End If
End Sub

Private Sub UnitModel_AfterUpdate()
        Me.Parent.ModelList.Requery
        Me.Parent.NoMatchModelsList.Requery
End Sub

Private Sub UnitModel_Change()
    If Me.Form.NewRecord = True Then
        Me.MakeID = Me.Parent.MakeList.Column(0)
        Me.UnitMake = Me.Parent.MakeList.Column(1)
        'Me.AddLabel.Requery
    End If
    If Nz(Me.Leased, 0) > 0 Or Nz(Me.Tracked, 0) > 0 Then
        MsgBox "You Cannot Change '" & Me.UnitModel & "'" & vbNewLine & "It is in use", vbCritical, "Type In Use"
        Me.Undo
       Exit Sub
    End If
End Sub

Private Sub SaveUpdate_Click()
    Me.Form.Refresh
   ' Me.AddLabel.Requery
    Me.Parent.ModelList.Requery
    Me.Parent.NoMatchModelsList.Requery
End Sub
