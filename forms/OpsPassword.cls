Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Password_AfterUpdate()
    Call SubmitRequest_Click
End Sub


Private Sub SubmitRequest_Click()
    If Me.Password = "LRFA" Or Me.Password = "1972" Or Me.Password = "CJFA" Or Me.Password = "PAFA" Or Me.Password = "MGFA" Or Me.Password = "ITFA" Or Me.Password = "EIFA" Then
        If Me.OpenArgs = "GroupForm" Then
            Form_GroupForm.LockUnLock (False)
            GoTo ExitMe
        Else
            Forms!Unitsform.UnitDocAccpt.Locked = False
        End If
    Else
        MsgBox "Incorrect Password", vbOKOnly, "Enter Password"
        If Me.OpenArgs = "GroupForm" Then
            Form_GroupForm.LockUnLock (True)
            GoTo ExitMe
        Else
            Forms!Unitsform.UnitDocAccpt.Locked = True
        End If
    End If
ExitMe:
    DoCmd.close acForm, "Opspassword", acSaveNo
End Sub

Private Sub CancelRequest_Click()

    DoCmd.close acForm, "Opspassword", acSaveNo

End Sub
