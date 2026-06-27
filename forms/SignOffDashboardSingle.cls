Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub CompanyName_Click()
    MessageUser "Open", "Opening Client's Form", "just a moment.."
    DoCmd.OpenForm "ClientsView", , , , , , Me.clientID
    MessageUser "Close"
End Sub

'Private Sub MasterLease_Click()
'    If Me.MLID & 0 <> 0 Then
'        MessageUser "Open", "Opening Master Lease Form", "just a moment.."
'        DoCmd.OpenForm "MstrLease", , , "MLID =" & Me.MLID
'        MessageUser "Close"
'        End If
'End Sub

'Private Sub ClientList_Click()
'    Dim SQLz   As String
'    SQLz = "SELECT * FROM vw_UnitGroupHealth "
'    SQLz = SQLz & "WHERE ClientID=" & Me.ClientList.Column(0) & " "
'    SQLz = SQLz & "AND [Assigned] <>'F' AND [Funded]='N'"
'    SQLz = SQLz & "OR (ClientID=" & Me.ClientList.Column(0) & " AND [Funded]='Y' AND [Assigned]='P' ) "
'    SQLz = SQLz & "ORDER BY SchID Desc;"
'    Me.Form.RecordSource = SQLz
'End Sub

Private Sub CloseMe_Click()
    DoCmd.OpenForm "FAMainMenu"
    DoCmd.close acForm, Me.Form.Name
End Sub



Private Sub OpenMassMod_Click()
    If Me.groupID & "test" = "test" Then
        MsgBox "Sorry, but you must first select a group", , "No Group, No Service :("
        Exit Sub
    End If
    MessageUser "Open", "Opening Mass-Mod Form", "just a moment.."
    DoCmd.close acForm, "BulkChange_form"
    DoCmd.OpenForm "BulkChange_form"
    Forms!BulkChange_Form.ClientList = DLookup("ClientID", "vw_SixKeys", "sGrpID=" & Me.groupID & "")
    Form_BulkChange_Form.ClientList_Click
    Forms!BulkChange_Form.LeaseList = DLookup("MLID", "vw_SixKeys", "sGrpID=" & Me.groupID & "")
    Form_BulkChange_Form.Leaselist_Click
    Forms!BulkChange_Form.ScheduleList = DLookup("SchID", "vw_SixKeys", "sGrpID=" & Me.groupID & "")
    Form_BulkChange_Form.ScheduleList_Click
    Forms!BulkChange_Form.GroupList = Me.groupID
    Form_BulkChange_Form.GroupList_Click
    Form_BulkChange_Form.GetAll_Click
     MessageUser "Close"
End Sub
Private Sub Purchasing_Click()
    Call StartSignOff(Me.groupID, "SignOff1")
    Me.Form.Requery
End Sub
Private Sub InvIn_Click()
    Call StartSignOff(Me.groupID, "SignOff2")
    Me.Form.Requery
End Sub

Private Sub MSOValid_Click()
    Call StartSignOff(Me.groupID, "SignOff3")
    Me.Form.Requery
End Sub
Private Sub Operations_Click()
    Call StartSignOff(Me.groupID, "SignOff4")
    Me.Form.Requery
End Sub

Private Sub TMT_Click()
    Call StartSignOff(Me.groupID, "SignOff5")
    Me.Form.Requery
End Sub

Private Sub Docs_Click()
    Call StartSignOff(Me.groupID, "SignOff6")
    Me.Form.Requery
End Sub
Private Sub Syndication_Click()
    Call StartSignOff(Me.groupID, "SignOff7")
    Me.Form.Requery
End Sub
Private Sub Schedule_DblClick(Cancel As Integer)
    If Me.groupID & "test" = "test" Then Exit Sub
    DoCmd.close acForm, "Schedule_Form"
    DoCmd.OpenForm "Schedule_Form", , , "SchID=" & DLookup("SchID", "vw_SixKeys", "sGrpID=" & Me.groupID & "")
End Sub
Private Sub UnitGroup_DblClick(Cancel As Integer)
    If Me.groupID & "test" = "test" Then
        MsgBox "Sorry, but you must first select a group", , "No Group, No Service :("
        Exit Sub
    End If
    DoCmd.close acForm, "GroupForm"
    DoCmd.OpenForm "GroupForm", , , "SGrpID=" & Me.groupID
End Sub
