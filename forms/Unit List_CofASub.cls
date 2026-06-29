Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz       As String
Private Sub AcceptDateSubmit_Click()
    On Error GoTo FixMe
    SQLz = "UPDATE Units SET UnitAcceptDate = #" & Me.AcceptDateCBO & "# WHERE (IsNull(UnitAcceptDate) or UnitAcceptDate <> #" & Me.AcceptDateCBO & "#) AND  GroupREF=" & Me.groupID & ";"
    DoCmd.RunSQL (SQLz)
    Call SetStatusA
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub COABackSubmit_Click()
    On Error GoTo FixMe
    SQLz = "UPDATE Units SET UnitCOABack = #" & Me.COABacKCBO & "# WHERE (IsNull(UnitCOABack) Or UnitCOABack <> #" & Me.COABacKCBO & "#) AND  GroupREF=" & Me.groupID & ";"
    DoCmd.RunSQL (SQLz)
    Call SetStatusA
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub CofACopySubmit_Click()
    On Error GoTo FixMe
    SQLz = "UPDATE Units SET UnitCofACopy = #" & Me.CofACopyCBO & "# WHERE (IsNull(UnitCofACopy) Or UnitCofACopy <> #" & Me.CofACopyCBO & "#) AND  GroupREF=" & Me.groupID & ";"
    DoCmd.RunSQL (SQLz)
    Call SetStatusA
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub
Private Sub SetStatusA()
On Error GoTo FixMe
    SQLz = "Update Units Set UnitStatus = 'A' Where (IsDate(UnitCOABack) OR IsDate(UnitCofACopy)) And UnitStatus <> 'A' AND  GroupREF=" & Me.groupID & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    If AllUnitsInGroup_Accepted(Me.groupID) = True Then Call AllUnitsInGroup_AcceptedMessage(Me.groupID)
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub
Private Sub Form_Current()
    Me.COABacKCBO = ""
    Me.CofACopyCBO = ""
    Me.AcceptDateCBO = ""
End Sub

Private Sub UnitAcceptDate_AfterUpdate()
    If IsDate(Me.UnitCOABack) And Me.UnitStatus = "B" Then Me.UnitStatus = "A"
    If AllUnitsInGroup_Accepted(Me.groupID) = True Then Call AllUnitsInGroup_AcceptedMessage(Me.groupID)
End Sub

Private Sub UnitCOABack_AfterUpdate()
    If IsDate(Me.UnitCOABack) And Me.UnitStatus = "B" Then Me.UnitStatus = "A"
    If AllUnitsInGroup_Accepted(Me.groupID) = True Then Call AllUnitsInGroup_AcceptedMessage(Me.groupID)
End Sub

Private Sub UnitCOABack_GotFocus()
    If Not IsDate(Me.UnitAcceptDate) Then
        MsgBox "Accepted date is REQUIRED.", vbCritical, "No Accepted Date"
        Me.UnitAcceptDate.SetFocus
    Else
    End If
End Sub

Private Sub UnitCofACopy_AfterUpdate()
    If IsDate(Me.UnitCofACopy) And Me.UnitStatus = "B" Then Me.UnitStatus = "A"
    If AllUnitsInGroup_Accepted(Me.groupID) = True Then Call AllUnitsInGroup_AcceptedMessage(Me.groupID)
End Sub
