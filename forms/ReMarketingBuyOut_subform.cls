Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public SQLz    As String
Public Sub ClearBoxes()
    Dim Ccrl   As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Medit" Then
            Me(Ccrl.Name).Value = ""
        End If
    Next Ccrl
End Sub
Private Sub BuyOutDate_AfterUpdate()
 '   Dim StatusX As String

 '   If IsDate(Me.BuyOutDate) Then
 '       If Me.unitstatus = "S" Then StatusX = "S" Else StatusX = "R"
 '       SQLz = "UPDATE Units SET UnitStatus =  '" & StatusX & "' " & _
 '              "WHERE UnitId=" & Me.UnitID & ";"
 '   Else
 '       Me.BuyOutDate = ""
 '       Me.unitstatus = "TR"
 '       SQLz = "UPDATE Units SET UnitStatus =  'TR' " & _
 '             "WHERE UnitId=" & Me.UnitID & ";"
 '   End If
 '   DoCmd.SetWarnings False
   ' DoCmd.RunSQL (SQLz)
 '   DoCmd.SetWarnings True
End Sub

Private Sub BuyOutCost_AfterUpdate()
    On Error GoTo FixMe
    If IsNull(Me.IFAmount) Then Me.IFAmount = Me.BuyOutCost
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub BuyOutDateSubmit_Click()
    Dim Response, SQLset As String
    SQLset = "SET BuyOutDate =#" & BuyOutDateCBO & "#;"
    If Me.BuyOutDate & "test" = "test" Then
        SQLset = "SET BuyOutDate = null;"
         Response = MsgBox("You're about to clear the Buy Out Dates, Are you sure?", vbQuestion + vbYesNo + vbCritical, "Clear Dates?")
        If Response = vbNo Then Exit Sub
    End If
    Call ConfrimIT("BuyOutDateCBO", "Buy Out Date")
    On Error GoTo FixMe
    SQLz = "UPDATE DISTINCTROW ReMarketingBuyOut_WorkTable INNER JOIN RemarketingUnitsBySchedule " & _
           "ON ReMarketingBuyOut_WorkTable.UnitID = RemarketingUnitsBySchedule.UnitID " & SQLset
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True

FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub
Private Sub BuyOutSubmit_Click()

    Dim Response, SQLset As String
    SQLset = ""
    If Me.BuyOutPrice = 0 Then
        SQLset = "SET BuyOutCost = 0;"
        Response = MsgBox("You're about to Zero ALL of the Buy Out Prices, Are you sure?", vbQuestion + vbYesNo + vbCritical, "Zero it?")
        If Response = vbNo Then Exit Sub
    End If
    If Me.BuyOutPrice Is Null Then
        SQLset = "SET BuyOutCost = NULL;"
        Response = MsgBox("You're about to NULL ALL of the Buy Out Prices, Are you sure?", vbQuestion + vbYesNo + vbCritical, "Nulls are scary?")
        If Response = vbNo Then Exit Sub
    End If

    If Nz(Me.BuyOutPrice, 0) > 0 Then SQLset = "SET BuyOutCost = " & Me.BuyOutPrice & " Where isnull(BuyOutCost) or BuyOutCost < 1;"

    Call ConfrimIT("BuyOutPrice", "Buy Out Cost")
    On Error GoTo FixMe
    SQLz = "UPDATE DISTINCTROW ReMarketingBuyOut_WorkTable INNER JOIN RemarketingUnitsBySchedule " & _
           "ON ReMarketingBuyOut_WorkTable.UnitID = RemarketingUnitsBySchedule.UnitID " & SQLset
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub Form_Current()
    Me.txtCurrent = Me.UnitID
End Sub

Private Sub IFAcctNumberSubmit_Click()
    Call ConfrimIT("IFAccountCBO", "Loan Number")
    On Error GoTo FixMe
    SQLz = "UPDATE DISTINCTROW ReMarketingBuyOut_WorkTable INNER JOIN RemarketingUnitsBySchedule " & _
           "ON ReMarketingBuyOut_WorkTable.UnitID = RemarketingUnitsBySchedule.UnitID " & _
           "SET IFLoanNumber = '" & Me.IFAccountCBO & "';"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub IfAmountSubmit_Click()
    Call ConfrimIT("IFAmountCbo", "Interim Amount")
    On Error GoTo FixMe
    SQLz = "UPDATE DISTINCTROW ReMarketingBuyOut_WorkTable INNER JOIN RemarketingUnitsBySchedule " & _
           "ON ReMarketingBuyOut_WorkTable.UnitID = RemarketingUnitsBySchedule.UnitID " & _
           "SET IFAmount = " & Me.IFAmountCbo & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub IfBankSubmit_Click()
    Call ConfrimIT("IFBank", "Interim Bank")
    On Error GoTo FixMe
    SQLz = "UPDATE DISTINCTROW ReMarketingBuyOut_WorkTable INNER JOIN RemarketingUnitsBySchedule " & _
           "ON ReMarketingBuyOut_WorkTable.UnitID = RemarketingUnitsBySchedule.UnitID " & _
           "SET IFBankID = '" & Me.IFBank & "';"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub IFinDateSubmit_Click()
    Call ConfrimIT("IFinDate", "Interim Date")
    On Error GoTo FixMe
    SQLz = "UPDATE DISTINCTROW ReMarketingBuyOut_WorkTable INNER JOIN RemarketingUnitsBySchedule " & _
           "ON ReMarketingBuyOut_WorkTable.UnitID = RemarketingUnitsBySchedule.UnitID " & _
           "SET IFDate = '" & Me.IFinDate & "';"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub PayOffDateSubmit_Click()
    On Error GoTo FixMe
    If Me.IFPaidDateCbo & "test" <> "test" Then
        MsgBox "You can not clear exisiting Paid Dates", vbCritical, "No Overwrite"
        Exit Sub
    End If
    Call ConfrimIT("IFPaidDateCbo", "Interim Funding Paid Date")
    SQLz = "UPDATE DISTINCTROW ReMarketingBuyOut_WorkTable INNER JOIN RemarketingUnitsBySchedule " & _
           "ON ReMarketingBuyOut_WorkTable.UnitID = RemarketingUnitsBySchedule.UnitID " & _
           "SET IFPaidDate = #" & Me.IFPaidDateCbo & "# Where isnull(IFPaidDate);"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
FixMe:
Me.Form.Refresh
Exit Sub
End Sub

Public Sub ConfrimIT(ByVal FieldBoX As String, ByVal TragetX As String)
    Dim Committ As String
    If Me(FieldBoX).Value & "TEST" = "TEST" Then
        Committ = MsgBox("You're about set the " & TragetX & " to Null, for ALL units listed." & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Clear " & TragetX & "?")
        If Committ = vbNo Then Exit Sub
    Else
        Committ = MsgBox("You're about set the " & TragetX & " to " & Me(FieldBoX).Value & ", for ALL units listed." & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Change " & TragetX & "?")
    End If
End Sub

Private Sub BuyDateEffectiveSubmit_Click()
    Call ConfrimIT("BuyDateEffectiveCbo", "Effective Buy Date")
    SQLz = "UPDATE DISTINCTROW ReMarketingBuyOut_WorkTable INNER JOIN RemarketingUnitsBySchedule " & _
           "ON ReMarketingBuyOut_WorkTable.UnitID = RemarketingUnitsBySchedule.UnitID " & _
           "SET BuyDateEffective = '" & Me.BuyDateEffectiveCbo & "';"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub Send2Excel_Click()
    Dim XLS As String
    XLS = "|D,HH|I,HH|F,BC|H,$E|J,BC|M,$E|Q,$E|"
    SendFormData2Excel Me, "Units", "Buy Out Info for " & Me.UnitCount & " Units", True, XLS, Me.UnitCount
End Sub

Private Sub InvoicePaidDate_AfterUpdate()
    If IsDate(InvoicePaidDate) And Me.unitstatus = "K" Then Me.unitstatus = "KS"
    If IsDate(InvoicePaidDate) And Me.unitstatus = "R" Then Me.unitstatus = "S"
    If IsDate(InvoicePaidDate) And Me.unitstatus Like "E*" Then Me.unitstatus = "S"
End Sub

Private Sub UnitID_DblClick(Cancel As Integer)
    Call OpenUnitsForm(Me.UnitID)
    'DoCmd.OpenForm "UnitsForm", , , "UNITID=" & Me.UnitID & ""
    ' Forms!UnitsForm.UnitsTabControl.Value = 6
    Me.UnitNum.SetFocus
End Sub

Private Sub UnitIFBank_AfterUpdate()
    If IsNull(Me.IFDate) Then Me.IFDate = Me.BuyOutDate
    If IsNull(Me.IFAmount) Then Me.IFAmount = Me.BuyOutCost
End Sub
