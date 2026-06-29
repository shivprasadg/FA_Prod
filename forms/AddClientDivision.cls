Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub AddClientDvision_Click()
            If DCount("ClientGroupID", "ClientDivisions", "ClientGroupID=" & Me.ClientGroupID & " AND DivisionName='" & Me.NewDivisionName & "'") > 0 Then
                MsgBox "There already is a Division named: " & vbNewLine & "'" & Me.NewDivisionName & "'" & vbNewLine & "Please change the name and resubmit.", , "ATLAAS Core"
                Exit Sub
            End If
            If IsNull(Me.ClientID) Or Me.NewDivisionName & "test" = "test" Then
                MsgBox "There is NO ClientID or Division Assigned" & vbNewLine & "Form Will Close (Nothing Saved)", , "ATLAAS Core"
                Call CloseMe_Click
                Exit Sub
            End If
            Dim SQLz As String
            SQLz = "INSERT INTO ClientDivisions (DivisionName,ClientGroupID, ClientID) SELECT '" & Me.NewDivisionName & "'," & Me.ClientGroupID & "," & Me.ClientID & ";"
            DoCmd.SetWarnings False
            DoCmd.RunSQL (SQLz)
            DoCmd.SetWarnings True
            Forms!ClientsView.Locations_sub.Form.DivisionList.Requery
            Forms!ClientsView.Locations_sub.Form.DivisionID.Requery
            Call CloseMe_Click
    End Sub

Private Sub CloseMe_Click()
    DoCmd.close acForm, Me.Form.Name
End Sub



Private Sub Form_Load()
 'If Me.OpenArgs <> "" Then
 '   Me.clientID = Split(Me.OpenArgs, "|", 0)
 '   Me.ClientGroupID = Split(Me.OpenArgs, "|", 1)
 'End If
End Sub
