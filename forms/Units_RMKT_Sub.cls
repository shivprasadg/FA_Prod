Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim iUserAccess As Integer
Dim iUserName, SQLz As String
Dim ValueX     As Double


Private Sub AdjBookValue_AfterUpdate()
   Call SetGainLoss
End Sub

Private Sub BuyOrdersList_DblClick(Cancel As Integer)
    If Me.BuyOrdersList & "test" = "test" Then Exit Sub
    Call OpenBuyersFrom(Me.BuyOrdersList.Column(3), Me.BuyOrdersList.Column(0))
End Sub

Private Sub BuyOutCost_AfterUpdate()
   Call SetGainLoss
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
  '  If Me.Dirty Then Me.UnitREF = Forms!Unitsform.Form.UnitID
End Sub

Private Sub Form_Close()
    Call RmktGainLossFix(True)
    On Error GoTo Form_Close_Error

    On Error GoTo 0
    Exit Sub

Form_Close_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Close of VBA Document Form_Units_RMKT_Sub"
End Sub

Private Sub SetLocks()
'Collect the User Name and Access Level
'Set Field level access rights
    On Error GoTo FixMe
    If FAUserAccessLevel < 10 Then Me.BuyOutCost.Visible = False Else Me.BuyOutCost.Visible = True
    If FAUserAccessLevel >= 6 Or FAUserAccessGroup Like "*ACCT*" Or FAUserAccessGroup Like "*OPS*" Then
        Me.BuyerRef.Locked = False
        Me.RmktBOSSent1.Locked = False
        Me.SoldPrice.Locked = False
        Me.FABrkComm.Locked = False
        Me.RmktAdjNote.Locked = False
        Me.RmktSPRAdj.Locked = False
        Me.BuyOutDate.Locked = False

        Me.AllowDeletions = True
        Me.AllowAdditions = True
        Me.AllowEdits = True
    Else
        Me.AllowEdits = False
        Me.AllowDeletions = False
        Me.AllowAdditions = False
        Me.RmktBOSSent1.Locked = True
    End If
    If FAUserAccessLevel >= 14 Then
        Me.TermPenalty.Visible = True
        Me.TermPenalty.Locked = False
    Else
        Me.TermPenalty.Visible = False
        Me.TermPenalty.Locked = True
    End If
    If FAUserAccessGroup Like "*ITT*" Or FAUserName = "lmanin" Or FAUserAccessGroup Like "*ACCT*" Then
        Me.UNSaleUnit.Visible = True
    Else
        Me.UNSaleUnit.Visible = False
    End If
    Exit Sub
FixMe:
    MsgBox "There was an Error in SetLocks Sub with Form: " & Me.Form.Name & " Form - Error Code:" & Err.Number, vbCritical, "ATLAAS Core Error"
    Resume Next
End Sub
Public Function LoadData(UnitIDx As Long)
    Call Me.RmktExpense.Form.LoadData(UnitIDx)
    On Error Resume Next
    Me.BuyOrdersList.Requery
    If Nz(Me.UnitREF, 0) = 0 Or Nz(Me.UnitREF, 0) <> UnitIDx Then
        If IsNull(Me.UnitREF) Or Me.Form.NewRecord = True Then Exit Function
        Call SetGainLoss
        Me.BuyOrdersList.Requery
        Call SetLocks
    End If
    On Error GoTo FixMe
    Call SetGainLoss
    Exit Function
FixMe:
    MsgBox "There was an Error in LoadData Sub with Form: " & Me.Form.Name & " Form - Error Code:" & Err.Number, vbCritical, "ATLAAS Core Error"
    Resume Next
End Function
Private Sub Form_Current()
    Me.BuyOrdersList.Requery
    If IsNull(Me.UnitREF) Or Me.Form.NewRecord = True Then Exit Sub
   ' Call SetGainLoss
End Sub
Public Sub SetGainLoss()
    Dim GLNumber As Double
    Dim ABVNumber As Double
    Dim BONumber As Double
    Dim NBONumber As Double
    If Me.Dirty = True Then Me.Dirty = False
    If Nz(Me.SoldPrice, 0) > 0 Then GLNumber = Me.SoldPrice - Nz(Me.AdjBookValue, 0)
    If IsDate(DLookup("unitoffleasedt", "Units", "UnitId=" & UnitREF)) Then
        NBONumber = Nz(DLookup("PDSurrenderAmt", "vw_PerDiemSurrender", "Not InspectionPassed Is Null and UnitId=" & Me.UnitREF), 0)
    End If
    BONumber = Nz(Me.BuyOutCost, 0)
    ABVNumber = Nz(Me.TermPenalty, 0)
    NBONumber = (BONumber - NBONumber)
    ABVNumber = (NBONumber - Abs(ABVNumber))
    If Nz(Me.GainLoss, 0) <> GLNumber Then Me.GainLoss = GLNumber
    If Nz(Me.BuyOutCost, 0) <> BONumber Then Me.BuyOutCost = BONumber
    If Nz(Me.BuyOutCostNet, 0) <> NBONumber Then Me.BuyOutCostNet = NBONumber
    If Nz(Me.AdjBookValue, 0) <> ABVNumber Then Me.AdjBookValue = ABVNumber

End Sub
Private Sub BuyerREF_DblClick(Cancel As Integer)
   ' If Me.BuyerRef & "test" <> "test" Then Call OpenBuyersFrom(Me.BuyerRef)
End Sub

Private Sub BuyOutDate_AfterUpdate()
    If Not IsNull(Me.BuyOutDate) And (Forms!Unitsform.[unitstatus] = "TR" Or Forms!Unitsform.[unitstatus] = "T") Then
        Forms!Unitsform.[unitstatus] = "R"
    Else
        Exit Sub
    End If
End Sub

Private Sub BuyOutDate_Dirty(Cancel As Integer)
    If Not IsNull(Me.BuyOutDate) And (Forms!Unitsform.[unitstatus] = "TR" Or Not IsNull(Me.BuyOutDate) And Forms!Unitsform.[unitstatus] = "T") Then
        Forms!Unitsform.[unitstatus] = "R"
    Else
        Exit Sub
    End If
End Sub

Private Sub InvoicePaidDate_AfterUpdate()
    If IsDate(Me.InvoicePaidDate) Then
        Me.Parent.unitstatus = "S"
    Else
        If Me.Parent.unitstatus = "S" Then Me.Parent.unitstatus = "R"
    End If
End Sub

Private Sub InvoiceSentDate_AfterUpdate()
    Call CheckSoldStatus
End Sub
Private Sub InvoiceInvoicePaidDate_AfterUpdate()
    Call CheckSoldStatus
End Sub
Private Sub CheckSoldStatus()
    If IsDate(Me.InvoiceSentDate) And Not IsDate(Me.InvoicePaidDate) Then Parent.unitstatus = "PS"
    If IsDate(Me.InvoiceSentDate) And IsDate(Me.InvoicePaidDate) Then Parent.unitstatus = "S"
    If Not IsDate(Me.InvoiceSentDate) And Not IsDate(Me.InvoicePaidDate) And IsDate(BuyOutDate) Then Parent.unitstatus = "R"
End Sub

Private Sub PakrDt_AfterUpdate()
    If Me.Parent.unitstatus = "S" Or Me.Parent.unitstatus = "PS" Or Me.Parent.unitstatus = "D" Or Me.Parent.unitstatus = "R" Then Exit Sub
    If IsDate(Me.PakrDt) Then
        MsgBox "This will Change the Unit's Status to 'P' (Parked)", vbInformation, "Unit Status Change"
        Me.Parent.unitstatus = "P"
        Exit Sub
    End If

End Sub

Private Sub SaveMe_Click()
    Call SetMousePointer(0)
    Me.Form.Refresh
End Sub

Private Sub SoldPrice_AfterUpdate()
    Call SetGainLoss
End Sub

Private Sub SUF_Returned_AfterUpdate()
   Call AutoUpdateEOL_Statuses(True)
End Sub

Private Sub SUF_Sent_AfterUpdate()
   Call AutoUpdateEOL_Statuses(True)
End Sub

Private Sub TermPenalty_AfterUpdate()
    If IsNull(Me.TermPenalty) Then Me.TermPenalty = 0
    'If Me.TermPenalty > 0 Then Me.TermPenalty = (Me.TermPenalty * -1)
    If ValueX <> 0 And Nz(Me.TermPenalty, 0) <> ValueX Then
        If ConfrimIT("TermPenalty", "Term Penalty") = False Then
            Me.TermPenalty = ValueX
            Exit Sub
        Else
            SQLz = "INSERT INTO ExitOptionsTermPenaltyChanges ( UnitID, UserID, ChangeDate, PrevValue, NewValue )  " & _
                   "SELECT " & Me.UnitREF & ", " & FAUserIDx & ", #" & Date & "#, " & ValueX & "," & Me.TermPenalty & ";"
            DoCmd.SetWarnings False
            DoCmd.RunSQL (SQLz)
            DoCmd.SetWarnings True
            Dim msg1 As String
            'make email message
            msg1 = msg1 & "<html><body><p style='color:RED;font-family:Segoe UI;font-size:20px;'><b>Termination Penalty Change</b></p>"
            msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:12px'>To: Jim Wharton<br />"
            msg1 = msg1 & "Subject:&#160;&#160;&#160;&#160;&#160;&#160;Termination Penalty Changed, Re AssetID:<b> " & Me.UnitREF & "</b><br />"
            msg1 = msg1 & "Priority:<b>&#160;&#160;&#160;&#160;&#160;&#160;High </b><br/><br/>"
            msg1 = msg1 & "Dear Jim,<br/><br/>There was a change submitted to a Remarketing Unit's Term-Penalty, the details are below.<br /><br />"
            msg1 = msg1 & "Company " & Me.Parent.UnitSouceInfo & "<br />"
            msg1 = msg1 & "AssetID: <B>" & Me.UnitREF & "</B><br />"
            msg1 = msg1 & "Prev Value: <B>" & Format(ValueX, "$#,#.#0") & "</B><br />"
            msg1 = msg1 & "New Value: <B>" & Format(Me.TermPenalty, "$#,#.#0") & "</B><br />"
            msg1 = msg1 & "User who made change: <B>" & FAUserFullName & "</B><br />"
            msg1 = msg1 & "<p style='font-family:Arial ; font-size:11px ; margin-left:30px'>" & Msgx & "</P>"
            msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:8px'><br/>[End Msg]"
            msg1 = msg1 & "Email Created: " & Date & " (via ATLAAS Core Data Manager).<br/><br />"
            msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:12px'><br/>"
            msg1 = msg1 & "Regards,<br /><br /><br />"
            msg1 = msg1 & "</body></html>"
            Call SendEmailMessage("jwharton@fleetadvantage.net", "Term Penatly Change", msg1, True, True, False, "ljohnson@fleetadvantage.net")
        End If
    End If
    Call SetGainLoss
End Sub

Public Sub TermPenalty_GotFocus()
    ValueX = Nz(Me.TermPenalty, 0)
End Sub
Public Function ConfrimIT(ByVal FieldBoX As String, ByVal TragetX As String) As Boolean
    Dim Committ As String
    ConfrimIT = True
    Committ = MsgBox("You're about Change the " & TragetX & " to " & Format(Me(FieldBoX).Value, "$#,#0") & ", from " & Format(ValueX, "$#,#0") & "." & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Change " & TragetX & "?")
    If Committ = vbNo Then ConfrimIT = False
End Function

Private Sub UNSaleUnit_Click()
    Call UnSellAUnit(Me.UnitREF)
    Me.Refresh
End Sub
