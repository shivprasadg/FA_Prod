Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz       As String

Private Sub ChildPartModel_Change()
' Me.UnitMake.Requery
'    Me.AddLabel1.Requery
    Me.AddLabel2.Requery
End Sub

Private Sub DeletedMe_Click()
    If Me.Form.NewRecord = True Then Exit Sub
    Dim Committ, Typex, Msgx As String
    Dim UnitX, Childx As Integer
    If Nz(Me.InUse, 0) > 0 Then
        MsgBox "You Cannot Delete '" & Me.ChildPartModel & "'" & vbNewLine & "It is in use", vbCritical, "Model In Use"
        Exit Sub
    End If
    Committ = MsgBox("You're about to Delete " & Me.ChildPartModel & vbNewLine & "Are you Sure?", vbQuestion + vbYesNo + vbCritical, "Delete!")
    If Committ = vbNo Then
        Exit Sub
    Else
        SQLz = "Delete * FROM LookUp_ChildPartModel WHERE ChildPartModelID = " & Me.ChildPartModelID & ""
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
        Me.Form.Requery
        Me.Parent.ModelList.Requery
        'Me.Parent.NoMatchModelsList.Requery
    End If
End Sub

Private Sub UnitModel_AfterUpdate()
    Me.Parent.ModelList.Requery
    Me.Parent.NoMatchModelsList.Requery
End Sub

Private Sub SaveUpdate_Click()
    Me.Form.Refresh
    Me.AddLabel1.Requery
    Me.AddLabel2.Requery
    Me.Parent.ModelList.Requery
    If DCount("ChildPartType", "LookUp_ChildPartType", "ChildPartType='" & Me.PartType & "'") < 1 Then
        Dim Committ As String
        Committ = MsgBox("Add A Part Type" & vbNewLine & "The Part Type: '" & Me.PartType & "'" & vbNewLine & _
                  "is NOT yet in ATLAAS Core" & vbNewLine & _
                  "Do You want to add it?", vbQuestion + vbYesNo + vbCritical, "Add Part Type!?")
        If Committ = vbNo Then
            Exit Sub
        Else
            DoCmd.SetWarnings False
            DoCmd.RunSQL ("Insert Into LookUp_ChildPartType (ChildPartType) Select '" & Me.PartType & "'")
            DoCmd.SetWarnings True
        End If

    End If

    '  Me.Parent.NoMatchModelsList.Requery
End Sub
