Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz, Response As String
Private Sub DeleteMe_Click()
If IsDate(Me.PaidDate) Then
   MsgBox "You Cannot Remove a SOLD Unit", vbCritical, "Wash your hands often"
   Exit Sub
End If
If IsDate(Me.Sent) Then
    Response = MsgBox("This Unit: " & Me.UnitID & " Was already Sent" & vbNewLine & " Do you want to proceed?" & vbNewLine & "You will be asked to comfirm, again.", vbQuestion + vbYesNo + vbCritical, "Think...")
If Response = vbNo Then Exit Sub
End If
    Response = MsgBox("You're about to Remove UnitID " & Me.UnitID & vbNewLine & "From Buyers Order: " & Me.BuyersOrderID & "." & vbNewLine & "Are you sure?", vbQuestion + vbYesNo + vbCritical, "What? Delete!")
    If Response = vbNo Then Exit Sub
    On Error GoTo FixMe
     DoCmd.SetWarnings False
      '  SQLz = "UPDATE Rmkt SET BuyerRef= Null, Rmkt.InvoiceRequestDate = Null, Rmkt.InvoiceRequestNum = Null, Rmkt.InvoiceSentDate = Null, Rmkt.RmkSm = Null WHERE UnitRef=" & Me.UnitID & ";"
      '  DoCmd.RunSQL (SQLz)
        SQLz = "Delete From RMKT_BuyersOrders Where BuyersOrderID=" & Me.BuyersOrderID & " and UnitID=" & Me.UnitID & ";"
        DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Me.Form.Requery
    Me.Parent.UnAssignedUnitsList.Requery
    Me.Parent.InvoiceRequestID.Requery
FixMe:
    Me.Form.Refresh
End Sub

Private Sub RemoveAllUnits_Click()
If IsDate(Me.Paid) Then
   MsgBox "You Cannot Remove SOLD Units", vbCritical, "Wash your hands often"
   Exit Sub
End If
If IsDate(Me.Sent) Then
    Response = MsgBox("These Units Were Sent" & vbNewLine & " Do you want to proceed?" & vbNewLine & "You will be asked to comfirm, again.", vbQuestion + vbYesNo + vbCritical, "take a moment...")
If Response = vbNo Then Exit Sub
End If
    On Error GoTo FixMe
Response = MsgBox("You're about to Remove all Units" & vbNewLine & "From Buyers Order: " & Me.BuyersOrderID & "." & vbNewLine & "Are you sure?", vbQuestion + vbYesNo + vbCritical, "What? Delete!")
If Response = vbNo Then Exit Sub
    DoCmd.SetWarnings False
        SQLz = "Delete From RMKT_BuyersOrders Where BuyersOrderID=" & Me.BuyersOrderID & ";"
        DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Me.Parent.UnAssignedUnitsList.Requery
    Me.Parent.InvoiceRequestID.Requery
FixMe:
    Me.Form.Requery
End Sub

Private Sub UnitID_DblClick(Cancel As Integer)
   Call OpenUnitsForm(Me.UnitID)
        Me.UnitVIN.SetFocus
End Sub
