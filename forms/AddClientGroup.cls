Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub AddClientButton_Click()
    If Me.NewClientGroup & Me.ExistingClientGroup & "test" = "test" Then
        MsgBox "Sorry, but you MUST select or enter a Client Group Name", vbCritical, "Missing Client Group"
        Exit Sub
    End If
    If Me.Password = "1972" Or Me.Password = "FA2015" Then
    Dim CGID As Integer
        If Me.ExistingClientGroup & "test" = "test" Then
            If DCount("ClientGroupID", "ClientGroups", "ClientGroupName='" & Me.NewClientGroup & "'") > 0 Then
                MsgBox "There already is a Client Group named: " & Me.NewClientGroup & vbNewLine & "Please change the name and resubmit.", , "ATLAAS Core"
                Exit Sub
            End If
            Dim SQLz As String
            SQLz = "INSERT INTO ClientGroups ( ClientGroupName, ParentGroupId )SELECT '" & Me.NewClientGroup & "'AS ClientGroupName, 0 AS Expr2;"
            DoCmd.SetWarnings False
            DoCmd.RunSQL (SQLz)
            DoCmd.SetWarnings True
            CGID = DLookup("ClientGroupID", "ClientGroups", "ClientGroupName='" & Me.NewClientGroup & "'")
            Else
            CGID = Me.ExistingClientGroup.Column(0)
        End If
'Exit Sub
        DoCmd.close acForm, "ClientsView"
        DoCmd.OpenForm "ClientsView"
        Forms!ClientsView.Form.SetFocus
        Forms!ClientsView.Form.AllowAdditions = True
        DoCmd.GotoRecord , , acNewRec
        Forms!ClientsView.Form.ClientGroupID = CGID
        Call CloseMe_Click
    Else
        MsgBox "Wrong Password... Try again.", , "Are you a Spy?"
        Me.Password = ""
        Me.Password.SetFocus
    End If
End Sub

Private Sub CloseMe_Click()
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub ExistingClientGroup_Click()
Me.NewClientGroup = Null
End Sub

Private Sub NewClientGroup_Change()
Me.ExistingClientGroup = Null
End Sub
