Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz As String

Private Sub Form_Close()
On Error Resume Next
    Dim FormX  As String
    If IsOpen("BuyersForm") = True Then Form_BuyersForm.Form.ProposalManager.Form.UpdateEvents
    If IsOpen("RemarketingSalesPersonsDashBoard") = True Then
        FormX = "RemarketingSalesPersonsDashBoard"
        Forms(FormX).proposalList.Requery
        Forms(FormX).EventHistory_Sub.Form.Requery
        Forms(FormX).EventHistory_Sub.Form.Refresh
        Form_RemarketingSalesPersonsDashBoard.GoRecord
        Form_RemarketingSalesPersonsDashBoard.Form.EventHistory_Sub.Form.Buyer_Click
    End If
    If IsOpen("EventPopUp") = True Then
        FormX = "EventPopUp"
        Forms(FormX).Form.Requery
        Forms(FormX).Form.Refresh
    End If
    If IsOpen("BuyerPopUp") = True Then
        FormX = "BuyerPopUp"
        Forms(FormX).EventHistory_Buyer.Form.Requery
        Forms(FormX).EventHistory_Buyer.Form.Refresh
    End If
On Error GoTo 0
End Sub

Private Sub Form_Current()
If Me.proposalStatus = "Closed" Then
Me.LockIt.caption = "UnLocked"
Call LockUnLockBtn_Click
End If
End Sub

Private Sub Form_Load()
    SQLz = "SELECT * FROM Buyer2Proposal_Query WHERE ProposalID = " & Me.proposalID & " ;"
    Me.proposalUnitsList.RowSource = SQLz
    ' Me.Insplink = ""
    SQLz = "Select * From Event_History Where ProposalID =" & Me.proposalID & " Order By EventStatus Desc, EventDueDate Desc  ;"
    Me.EventHistory_proposal.Form.RecordSource = SQLz
    Me.EventHistory_proposal.Form.Refresh
End Sub

Private Sub LockUnLockBtn_Click()
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "EditYN" Then
            If Me.LockIt.caption = "Locked" Then Me(Ccrl.Name).Locked = False Else Me(Ccrl.Name).Locked = True
        End If
    Next Ccrl
    If Me.LockIt.caption = "Locked" Then
        Me.LockIt.caption = "Unlocked"
        Me.LockUnLockBtn.Picture = "C:\FAAS\Images\UnLocked.png"
    Else
        Me.LockIt.caption = "Locked"
        Me.LockUnLockBtn.Picture = "C:\FAAS\Images\Locked.png"
    End If
End Sub

Private Sub ProposalUnitsList_DblClick(Cancel As Integer)
    If IsNull(Me.proposalUnitsList.Column(0)) Then Exit Sub
    Call OpenUnitsForm(Me.proposalUnitsList.Column(0))
    'DoCmd.OpenForm "UnitsForm", , , "UnitID=" & Me.proposalUnitsList.Column(0)
    Pause (0.5)
    DoCmd.SelectObject acForm, "UnitsForm"
    Forms!Unitsform.UnitsTabControl.Value = 6
End Sub

Private Sub SaveUpdate_Click()
DoCmd.Save acForm, Me.Form.Name
Form_Close
End Sub
