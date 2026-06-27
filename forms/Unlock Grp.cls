Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

'------------------------------------------------------------
' unlock_Click
'
'------------------------------------------------------------
Private Sub unlock_Click()

    Dim iUserAccess As Integer
        iUserAccess = DLookup("[Access Level]", "TBL_Employees", "[User ID] = FAUserName()")

         If iUserAccess = 9 Or iUserAccess = 7 Or iUserAccess = 5 Or iUserAccess = 0 Then

          Forms.GroupForm.AllowAdditions = True
          Forms.GroupForm.AllowEdits = True
          Forms.GroupForm.AllowDeletions = True

    MsgBox "Form already unlocked"

    Exit Sub
    Else


    If (Forms![unlock grp]!PWord = "LRFA" Or Forms![unlock grp]!PWord = "IKFA" Or Forms![unlock grp]!PWord = "KIFA") Then
        Forms!GroupForm.AllowEdits = -1




    End If
End If
End Sub


'------------------------------------------------------------
' CLGrpULBtn_Click
'
'------------------------------------------------------------
Private Sub CLGrpULBtn_Click()

    DoCmd.close , ""

End Sub
