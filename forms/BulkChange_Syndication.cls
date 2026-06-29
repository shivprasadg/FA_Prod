Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz As String

Private Sub DrawDownId_AfterUpdate()
If Me.DrawDownId & "test" <> "test" Then
   Me.UnitIFAmt = Nz(Me.VendorCost, Me.UnitFinalCost)
Else
   Me.UnitIFAmt = Null
End If
End Sub

Private Sub DrawDownSubmit_Click()
    On Error GoTo FixMe
    If Nz(Me.DrawDownCbo, 1) = 1 Then
        Dim Committ As String
        Committ = MsgBox("You're about to Remove Units From a Draw" & vbNewLine & "Are you Sure?", vbQuestion + vbYesNo + vbCritical, "Remove Draw!")
        If Committ = vbNo Then Exit Sub
    End If
    Dim ExtraFields As String
    If Nz(Me.DrawDownCbo, 1) > 1 Then ExtraFields = ", Units.UnitIFAmt = Nz(VendorCost, UnitFinalCost)" Else ExtraFields = ", Units.UnitIFBankID = NULL, Units.UnitIFAmt = NULL"
    Call UpdateMassUnits("Units.DrawDownId", Me.DrawDownCbo, , ExtraFields)
FixMe:
    Me.Form.Refresh
End Sub

Private Sub Form_AfterUpdate()
   Me.Parent.ModMade = -1
End Sub

Private Sub Form_ApplyFilter(Cancel As Integer, ApplyType As Integer)
    Dim zFilter As String
    zFilter = Replace(Replace(Replace(Me.Form.Filter, "[Bulk_Update_SelectedUnits].", ""), "]", ""), "[", "")
    If Me.Form.Filter Like "*Bulk_Update_SelectedUnits_SYND*" Then _
       zFilter = Replace(Replace(Replace(Me.Form.Filter, "[Bulk_Update_SelectedUnits_SYND].", ""), "]", ""), "[", "")
    If Me.Form.Filter Like "*Draw*" Then zFilter = "DrawDownID=" & Me.DrawDownId
    ' End If
    If zFilter <> "" Then Call Me.Parent.SetUnits(zFilter)
    Call Me.Parent.SetOrderBy(Me.OrderBy)
    Me.Form.Filter = ""

End Sub
Private Sub Form_Current()
On Error GoTo whoops
   Me.txtCurrent = Me.UnitID
whoops:
End Sub

Public Sub ClearBoxes()
Dim Ccrl As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Medit" Then
            Me(Ccrl.Name).Value = ""
        End If
     Next Ccrl
    Me.DrawDownCbo.Requery
    Me.DrawDownId.Requery
End Sub

Private Sub Form_Dirty(Cancel As Integer)
   Me.Parent.ModMade = -1
End Sub

Private Sub IfAmountSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitIFAmt", "Nz(Units.VendorCost,Units.UnitFinalCost)")
FixMe:
    Me.Form.Refresh
End Sub

Private Sub IfBankSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitIFBankID", Me.IFBankCBO)
    Call SetDrawIFBank(Me.IFBankCBO)
    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub

Private Sub IFDateSubmit_Click()
    If Not FAUserAccessGroup Like "*SYND*" Then
        Call MsgBox("Sorry only the Syndication Dept" & vbNewLine & "can set the Draw Fund Date", vbCritical, "Don't even rhink about it!")
        Exit Sub
    End If
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitIFDate", Me.IFDate)
    Call SetDrawIFDate(Me.IFDate)
    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub

Private Sub PayOffDateSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitPayOff", Me.PayOffDate)
    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub

Private Sub RVGuarOEMSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.RVGuarOEM", Me.RVGuarOEMCurr)
    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub

Private Sub RVGuarSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.RVGuar", Me.RVGuarCurr)
    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub

Private Sub SyndRentSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UNITRENT", Me.SyndRentCurr)
    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub

Private Sub SynResidSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitSynResid", Me.SynResidCbo)
    ModMade = -1
FixMe:
    Me.Form.Refresh
End Sub

Private Sub UnitID_DblClick(Cancel As Integer)
Call OpenUnitsForm(Me.UnitID)
Me.unitvin.SetFocus
End Sub


Private Sub UnitIFAmt_AfterUpdate()
    If Me.UnitIFAmt > Me.VendorCost Then
        Call MsgBox("Sorry, you cannot enter a Value greater then the Unit Cost." & vbNewLine & "This will be undone", vbCritical, "Stop spending money you don't have!")
        Me.Undo
    End If
End Sub

Private Sub UnitIFBank_AfterUpdate()
    If Me.DrawDownId & "test" <> "test" Then Call SetInterimBank(Me.DrawDownId, Nz(Me.UnitIFBank, 0))
    On Error GoTo FixMe
    DoCmd.SetWarnings False
    Forms!BulkChange_Form.BulkChange_Syndication.Form.Requery
    Forms!BulkChange_Form.BulkChange_ChildParts.Form.Requery
    DoCmd.SetWarnings True
    Exit Sub
FixMe:
    Forms!BulkChange_Form.BulkChange_Syndication.Form.Requery
    Forms!BulkChange_Form.BulkChange_ChildParts.Form.Requery
    DoCmd.SetWarnings True
End Sub

Private Sub UnitIFDate_AfterUpdate()
   If Me.DrawDownId & "test" <> "test" Then Call SetInterimFundDates(Me.DrawDownId, Nz(Me.UnitIFDate, ""))
   DoCmd.SetWarnings False
   On Error GoTo FixMe
    DoCmd.SetWarnings False
    Forms!BulkChange_Form.BulkChange_Syndication.Form.Requery
    Forms!BulkChange_Form.BulkChange_ChildParts.Form.Requery
    DoCmd.SetWarnings True
    Exit Sub
FixMe:
    Forms!BulkChange_Form.BulkChange_Syndication.Form.Requery
    Forms!BulkChange_Form.BulkChange_ChildParts.Form.Requery
    DoCmd.SetWarnings True
End Sub

Private Sub UnitIFDate_BeforeUpdate(Cancel As Integer)
    If Not FAUserAccessGroup Like "*SYND*" Then
        Call MsgBox("Sorry only the Syndication Dept" & vbNewLine & "can set the Draw Fund Date", vbCritical, "Don't even rhink about it!")
        Me.Undo
    End If
End Sub
