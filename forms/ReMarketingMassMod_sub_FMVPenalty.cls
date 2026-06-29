Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public SQLz, SQLset As String
Public Committ As Variant
Dim RST        As DAO.Recordset
Public Sub ClearBoxes()
    Dim Ccrl   As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Medit" Then
            Me(Ccrl.Name).Value = ""
        End If
    Next Ccrl
End Sub

Private Sub AdjBookValue_AfterUpdate()
    Me.ProjPenalty.Requery
End Sub

Private Sub Form_Current()
    On Error Resume Next
    Me.txtCurrent = Me.UnitID
    SQLz = "Select ClientShortname As Client, Schedule, UnitGroup From vw_SixKeys WHERE AssetID=" & Me.UnitID & ";"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    Me.Parent.BuyNo = "Client: " & RST!Client & " Sch|Grp: " & RST!Schedule & "|" & RST!UnitGroup
    RST.close
    Set RST = Nothing
    On Error GoTo 0
End Sub

Private Sub ProjectedLPAFLossSubmit_Click()
    If Me.ProjectedLPAFLossCBO & "test" = "test" Then Exit Sub
    On Error GoTo FixMe
    Call UpdateRMKTtbl("RMKT.TermPenalty", Me.ProjectedLPAFLossCBO, , , Me.Parent.Name)
    Me.Form.Refresh
FixMe:
    Exit Sub
End Sub

Private Sub TargetSaleSubmit_Click()
    If TargetSaleCBO & "test" = "test" Then Exit Sub
    On Error GoTo FixMe
    Call UpdateRMKTtbl("RMKT.TargetSale", Me.TargetSaleCBO, , , Me.Parent.Name)
    Me.Form.Refresh
FixMe:
    Exit Sub
End Sub
Private Sub DiscountSubmit_Click()
    If Me.DiscountCBO & "test" = "test" Then Exit Sub
    On Error GoTo FixMe
    Call UpdateRMKTtbl("RMKT.Discount", Me.DiscountCBO, , , Me.Parent.Name)
    Me.Form.Refresh
FixMe:
    Exit Sub
End Sub

Private Sub FABrkCommSubmit_Click()
    If Me.FABrkCommCbo & "test" = "test" Then Exit Sub
    On Error GoTo FixMe
    Call UpdateRMKTtbl("RMKT.FABrkComm", Me.FABrkCommCbo, , , Me.Parent.Name)
    Me.Form.Refresh
FixMe:
    Exit Sub
End Sub

Private Sub CopySUFmileage_Click()
    Dim CopyIT As String
    CopyIT = MsgBox("You're about copy the PARKED Mileage to INSPECTED mileage" & vbNewLine & "(only for units missing INSPECTED mileage)" & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Copy SUF Mileage?")
    If CopyIT = vbNo Then Exit Sub
    SQLz = "UPDATE DISTINCTROW RMKT_MMod_Work INNER JOIN RemarketingUnitsBySchedule " & _
           "ON RMKT_MMod_Work.UnitID = RemarketingUnitsBySchedule.UnitID " & _
           "SET InspMileage1 = MiParked WHERE isnull(InspMileage1);"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
FixMe:
    Exit Sub
End Sub

Private Sub WhSValueSubmit_Click()
    If Me.WhSValueCBO & "test" = "test" Then Exit Sub
    On Error GoTo FixMe
    Call UpdateRMKTtbl("RMKT.WhSValue", Me.WhSValueCBO, , , Me.Parent.Name)
    Me.Form.Refresh
FixMe:
    Exit Sub
End Sub


Private Sub CopyBOSubmit_Click()
    Dim CopyIT As String
    CopyIT = MsgBox("You're about copy the Buy-Out Values to Book Values" & vbNewLine & "(only for units without a Book Value)" & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Copy Buy-Out as Book?")
    If CopyIT = vbNo Then Exit Sub
        SQLz = "UPDATE DISTINCTROW RMKT_MMod_Work INNER JOIN RemarketingUnitsBySchedule " & _
           "ON RMKT_MMod_Work.UnitID = RemarketingUnitsBySchedule.UnitID " & _
           "SET AdjBookValue = BuyOutCost WHERE BuyOutCost >0 and AdjBookValue <=0;"
    DoCmd.SetWarnings False
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub
Private Sub TermPenaltyCalculate_Click()
    Dim CopyIT As String
    CopyIT = MsgBox("You're about Calculate Term Penalty based on Buy Out, Book and Sold Values" & vbNewLine & "(for all listed units)" & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Calculate Term Penalty?")
    If CopyIT = vbNo Then Exit Sub
    SQLz = "UPDATE DISTINCTROW RMKT_MMod_Work INNER JOIN RemarketingUnitsBySchedule " & _
           "ON RMKT_MMod_Work.UnitID = RemarketingUnitsBySchedule.UnitID " & _
           "SET AdjBookValue = BuyOutCost WHERE BuyOutCost >0 and AdjBookValue <>0;"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    SQLz = "UPDATE DISTINCTROW RMKT_MMod_Work INNER JOIN RemarketingUnitsBySchedule " & _
           "ON RMKT_MMod_Work.UnitID = RemarketingUnitsBySchedule.UnitID " & _
           "SET GainLoss = SoldPrice - AdjBookValue WHERE BuyOutCost >0 and SoldPrice >0;"
    DoCmd.RunSQL (SQLz)
    SQLz = "UPDATE DISTINCTROW RMKT_MMod_Work INNER JOIN RemarketingUnitsBySchedule " & _
           "ON RMKT_MMod_Work.UnitID = RemarketingUnitsBySchedule.UnitID " & _
           "SET GainLoss = AdjBookValue - BuyOutCost WHERE BuyOutCost >0 and SoldPrice <>0;"
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Public Sub ConfrimIT(ByVal FieldBoX As String, ByVal TragetX As String)
    If Me(FieldBoX).Value & "TEST" = "TEST" Then
        Committ = MsgBox("You're about set the " & TragetX & " to Null, for ALL units listed." & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Update-Change confrim " & TragetX & "?")
        If Committ = vbNo Then Exit Sub
    Else
        Committ = MsgBox("You're about set the " & TragetX & " to " & Me(FieldBoX).Value & ", for ALL units listed." & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Change " & TragetX & "?")
    End If
    If Committ = vbNo Then Exit Sub
End Sub

Private Sub Send2Excel_Click()
    Dim XLS As String
    XLS = "|D,HH|I,HH|F,BC|H,$E|J,BC|M,$E|Q,$E|"
    SendFormData2Excel Me, "Units", "Buy Out Info for " & Me.UnitCount & " Units", True, XLS, Me.UnitCount
End Sub

Private Sub UnitID_DblClick(Cancel As Integer)
    Call OpenUnitsForm(Me.UnitID)
    Forms!Unitsform.UnitsTabControl.Value = 6
    Me.UnitNum.SetFocus
End Sub
