Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz As String

Private Sub Form_Current()
On Error GoTo whoops
   Me.txtCurrent = Me.UnitID
whoops:
   Exit Sub
End Sub
Private Sub Form_ApplyFilter(Cancel As Integer, ApplyType As Integer)
    Dim zFilter As String
    zFilter = Replace(Replace(Replace(Me.Form.Filter, "[Bulk_Update_SelectedUnits].", ""), "]", ""), "[", "")
    If zFilter <> "" Then Call Me.Parent.SetUnits(zFilter)
    Call Me.Parent.SetOrderBy(Me.OrderBy)
    Me.Form.Filter = ""
End Sub
Private Sub GoHome_Click()
DoCmd.OpenForm " FAMainMenu"
End Sub
Public Sub ClearBoxes()
Dim Ccrl As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Medit" Then
            Me(Ccrl.Name).Value = ""
        End If
     Next Ccrl
End Sub

Private Sub IfAmountSubmit_Click()
Exit Sub 'Disabled maual updates to IF amount 10/24/2025 Leo
    On Error GoTo FixMe
        Call UpdateMassUnits("Units.UnitIFAmt", "Nz(Units.VendorCost,Units.UnitFinalCost")
FixMe:
    Me.Form.Refresh
    Exit Sub

End Sub

Private Sub IfBankSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitIFBankID", Me.IFBankCBO)
FixMe:
    Me.Form.Refresh
    Exit Sub

End Sub

Private Sub IFDateSubmit_Click()
    On Error GoTo FixMe
    If FAUserAccessGroup Like "*SYND*" Then
        Call UpdateMassUnits("Units.UnitIFDate", Me.IFDate)
    Else
        Call MsgBox("Sorry only the Syndication Dept" & vbNewLine & "can set the Draw Fund Date", vbCritical, "Don't even think about it!")
    End If
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub PayOffDateSubmit_Click()
    On Error GoTo FixMe
    Call UpdateMassUnits("Units.UnitPayOff", Me.PayOffDate)
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub UnitID_DblClick(Cancel As Integer)
Call OpenUnitsForm(Me.UnitID)
Me.UnitUnitNum.SetFocus
End Sub

Private Sub UnitIFAmt_AfterUpdate()
    If Me.UnitIFAmt > Me.VendorCost Then
        Call MsgBox("Sorry, you cannot enter a Value greater then the Vendor Cost." & vbNewLine & "This will be undone", vbCritical, "Stop spending money you don't have!")
        Me.Undo
    End If
End Sub

Private Sub UnitIFBank_AfterUpdate()
      If Me.DrawDownId & "test" <> "test" Then Call SetInterimBank(Me.DrawDownId, Nz(Me.UnitIFBank, 0))
End Sub

Private Sub UnitIFDate_AfterUpdate()
    If Me.DrawDownId & "test" <> "test" Then Call SetInterimFundDates(Me.DrawDownId, Nz(Me.UnitIFDate, ""))
End Sub

Private Sub UnitIFDate_BeforeUpdate(Cancel As Integer)
    If Not FAUserAccessGroup Like "*SYND*" Then
        Call MsgBox("Sorry only the Syndication Dept" & vbNewLine & "can set the Draw Fund Date", vbCritical, "Don't even rhink about it!")
        Me.Undo
    End If
End Sub
