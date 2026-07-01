Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public SQLz, ShowFee, ValueX, TickedX, Msgx, PrevX As String
Public Schx As Integer
Public RST As DAO.Recordset
Public CurrentIndicator As String
Dim Committ As String



Private Sub AddDocLink_Click()
    DoCmd.close acForm, "DocAddEditLinkForm"
    DoCmd.OpenForm "DocAddEditLinkForm", acNormal, , , acFormAdd, , Me.ClientID & "|" & Me.MLID & "|" & Me.SchID & "|" & Me.GroupID & "|0|" & Me.Form.Name & "|Add"
End Sub

Private Sub AssigneeBankID_Click()
    Me.Assignee_Bank = Me.AssigneeBankID.Column(1)
End Sub

Private Sub CloneSchGroup_Click()
    If FAUserAccessGroup Like "*ITT*" Or FAUserAccessGroup Like "*TMT*" Or FAUserAccessGroup Like "*OPS*" Then
        DoCmd.close acForm, "CloneScheduleGroups"
        DoCmd.OpenForm "CloneScheduleGroups", , , , , , Me.SchID & "|" & Me.GroupID
    End If
End Sub

Private Sub CopyThisGroup_Click()
Exit Sub
    If FAUserAccessGroup Like "*ITT*" Or FAUserAccessGroup Like "*TMT*" Or FAUserAccessGroup Like "*OPS*" Then
        DoCmd.close acForm, "GroupAdd_Form"
        DoCmd.OpenForm "GroupAdd_Form", , , , , , Me.SchID & "|" & Me.GroupID
    End If
End Sub

Private Sub ExitOptions_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call MakeExitOptionLocal
End Sub

Private Sub Form_Load()

    Schx = Nz(Me.OpenArgs, Me.SchID)
    If FAUserAccessGroup Like "*ACCT*" Or FAUserAccessGroup Like "*SYND*" Or FAUserAccess > 14 Then
        Me.AssignedIndicator.Locked = False
        Me.FinalEconomicsReportBtn.Enabled = True
        Dim Ccrl As Control
        For Each Ccrl In Me.Controls
            If Me(Ccrl.Name).Tag = "Acct" Then
                Me(Ccrl.Name).Visible = True
            End If
        Next Ccrl
    Else
        Me.AssignedIndicator.Locked = True
        Me.FinalEconomicsReportBtn.Enabled = False
    End If
    If FAUserAccessGroup Like "*BDEX*" Then Me.TabControl.Pages(5).Visible = False
    Call HideShowUAT
    If Nz(Me.Assigned_Indicator, "P") = "F" Then Call LockUnLock(True) Else Call LockUnLock(False)
    Me.GroupsList = Me.SGrpID
    Me.DocLinksSub.Form.RecordSource = "Select * From DocsLinkSubQ WHERE SchID=" & Me.SchID
FixMe:
End Sub
Private Sub Form_Current()
    'On Error GoTo FixMe
    Me.AssignedIndicator = Me.Assigned_Indicator
    Me.ShowFee = "Show"
    Call CheckDocsFields("ResidualGuaranteeDocYN", "RGD", True)
    Call CheckDocsFields("ShortFallAgreementYN", "SFA", True)
    Call CheckDocsFields("IndemnificationDocYN", "IDF", True)
    Call CheckDocsFields("CustomDocumentNameYN", "IDF", True)
    Me.DocLinksSub.Requery
    Me.CofaSign.Requery
    If Nz(Me.LeaseExtOption1LRF.Value, 0) = 0 And Me.UnitsInGroup > 0 Then Me.LeaseExtOption1Rent = 0 Else Me.LeaseExtOption1Rent = (Me.OrigEquipCost / UnitsInGroup) * (LeaseExtOption1LRF)
    Dim Rst1 As DAO.Recordset
    If Me.MLOrig = "FA" Then
        Set Rst1 = CurrentDb.OpenRecordset("Select * From Group_OEC_Rent_LRF WHERE UnitGroupId= " & Me.SGrpID & ";")
        Me.CalcRent = CCur(Nz(Rst1!GroupRent, 0))
        Me.CalcSyndRV = CCur(Nz(Rst1!GroupSyndRV, 0))
        Me.CalcSalesRent = CCur(Nz(Rst1!GroupSalesRent, 0))
        Me.CalcSalesRV = CCur(Nz(Rst1!GroupSalesRV, 0))
        Me.CalcOEC = CCur(Nz(Rst1!GroupOEC, 0))
        Me.CalcVendorCost = CCur(Nz(Rst1!VendorPlus, 0))
        Me.ExtRollUp = CCur(Nz(Rst1!GroupExtRent, 0))
        Me.LeaseExtOption1LRF = Nz(Rst1!EXTLRF, 0)
        Me.MTMRollup = CCur(Nz(Rst1!GroupMTMRent, 0))
        Me.MTMRate = Nz(Rst1!MTMLRF, 0)
        If Me.LRFGroup & "test" = "test" Then Me.LRFGroup = CCur(Nz(Rst1!GroupLRF, 0))
        'Me.CalcGroupLRF = CCur(Nz(Rst!GroupLRF, 0))
        If Me.CalcRent = 0 Or Me.CalcOEC = 0 Then CalcGroupLRF = 0 Else CalcGroupLRF = CDbl(Me.CalcRent / Me.CalcOEC)
        If Me.LORebate.Value > 0 And Nz(Me.SynRebate.Value, 0) = 0 Then Me.SynRebate.BackColor = RGB(255, 255, 153) Else Me.SynRebate.BackColor = RGB(255, 255, 255)

        Rst1.close
        Set Rst1 = Nothing
    End If
    Me.BillToAddressID.Requery
    Me.BillingAddressPreview = Me.BillToAddressID.Column(1)
    Me.GroupRevisionLog.Requery
    Call CheckPerdiem
    If Me.OrigEquipCost & "test" = "test" And Me.MLOrig = "FA" Then Me.OrigEquipCost = Me.CalcOEC
    If Me.LeasePymtForGroup & "test" = "test" And Me.MLOrig = "FA" And Me.CalcSalesRent > 1 Then Me.LeasePymtForGroup = Me.CalcSalesRent
    '    If Me.LeasePymtForGroup & "test" = "test" And Me.MLOrig = "FA" And Me.CalcRent < 1 Then Me.LeasePymtForGroup = Me.CalcRent


    If Me.CalcOEC <> Me.OrigEquipCost And Me.MLOrig = "FA" Then
        Beep
        Me.OrigEquipCost = Me.CalcOEC
        MsgBox "The Group OEC was corrected to " & Format(Me.CalcOEC, "$#,#.#0") & vbNewLine & "Which is the actual 'roll-up' of the Unit Level cost.", , AtlaasVersion
        'Me.OrigEquipCost.BackColor = RGB(255, 204, 0)
        'Me.calcOEC.BackColor = RGB(255, 204, 0)
    Else
        'Me.OrigEquipCost.BackColor = RGB(255, 255, 255)
        'Me.calcOEC.BackColor = RGB(255, 255, 255)
    End If
    If Me.CalcSalesRent <> Me.LeasePymtForGroup And Me.MLOrig = "FA" And Me.AssignedIndicator <> "F" And Nz(Me.CalcRent, 0) = 0 Then
        Me.LeasePymtForGroup = Me.CalcSalesRent
        MsgBox "The Group Rent was corrected to " & Format(Me.CalcSalesRent, "$#,#.#0") & vbNewLine & "Which is the actual 'roll-up' of the Unit Sales Rent.", , AtlaasVersion
    End If
    If Me.CalcRent <> Me.LeasePymtForGroup Then
        Me.LeasePymtForGroup.BackColor = RGB(255, 204, 0)
        Me.CalcRent.BackColor = RGB(255, 204, 0)
    Else
        Me.LeasePymtForGroup.BackColor = RGB(255, 255, 255)
        Me.CalcRent.BackColor = RGB(255, 255, 255)
    End If

    ' If Me.SGrpFSLPmt > 0 Then Me.PVFeesYN = True
    If FAUserAccessLevel < 15 Then
        Me.ExtRollUp.Visible = False
        Me.MTMRollup.Visible = False
    End If
    If FAUserAccessLevel > 24 Then
        Me.SGrpNO.Visible = True
        If Me.Assigned_Indicator <> "F" Then Me.SGrpNO.Locked = False
    End If
    Call MakeWorkTable
    '  Call CheckFunded  'modified as per Laurie - 8/25/2016 lbj
    If UseSpecialRemitGrp = True Then Me.SpecialRemitAddrGrp.Locked = False
    SQLz = "SELECT * FROM ExitOptionsGroup INNER JOIN UnitGroup_WorkTable ON ExitOptionsGroup.AssetID = UnitGroup_WorkTable.UnitId; "    'WHERE GroupID=" & Me.SGrpID & ";"
    'Me.ExitOptionsGroup_Sub.Form.RecordSource = SQLz
FixMe:
End Sub
Private Sub MakeWorkTable()
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete * From UnitGroup_WorkTable")
    DoCmd.RunSQL ("Insert into UnitGroup_WorkTable (UnitID, GroupID) Select UnitID, GroupRef From Units Where GroupRef =" & Me.GroupID & "")
    DoCmd.SetWarnings True
End Sub
Private Sub AddChildPart_Click()
    Call OpenMassMod_Click
End Sub

Private Sub AddParts_Click()
    DoCmd.close acForm, "ChildPartsToGroupPO"
    DoCmd.OpenForm "ChildPartsToGroupPO", acNormal, , , , , Me.GroupID & "|" & Me.PONum1
End Sub

Private Sub AssignmentDate___AfterUpdate()
    Call CheckFunded
End Sub

Private Sub BillToAddressID_Click()
    Me.BillingAddressPreview = Me.BillToAddressID.Column(2)
End Sub


Private Sub CloseMe_Click()
    On Error Resume Next
    Forms!GroupForm.AllowEdits = -1
    If (IsNull(Forms!GroupForm!BLCD) And Forms!GroupForm!Assigned_Indicator = "f") Then
        Beep
        MsgBox "Assignee Ind of F requires BLCD to be set", vbCritical, "Assignee Indicator Error"
    End If
    Call CheckDocsFields("ResidualGuaranteeDocYN", "RGD", True)
    If TickedX = "Stop" Then Exit Sub
    Call CheckDocsFields("ShortFallAgreementYN", "SFA", True)
    If TickedX = "Stop" Then Exit Sub
    Call CheckDocsFields("IndemnificationDocYN", "IDF", True)
    If TickedX = "Stop" Then Exit Sub
    Call CheckFunded
    Call MakePaymentAdjustment(Me.GroupID)
    DoCmd.close acForm, Me.Form.Name
  End Sub

Private Sub CofaSign_Click()
    On Error GoTo FixMe

    Committ = MsgBox("You're about to set " & Me.CofaSign.Column(1) & " as the CofA Signer" & vbNewLine & "for all (" & Me.UnitsInGroup & ") units in this group" & vbNewLine & "Are you sure?", vbQuestion + vbYesNo + vbCritical, "Set CofA Signer")
    If Committ = vbNo Then
        Exit Sub
    Else
        SQLz = "UPDATE Units SET Units.CofASigner = " & Me.CofaSign.Column(0) & " WHERE (((Units.GroupREF)=" & Me.GroupID & "));"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
    End If
    Exit Sub
FixMe:
    MsgBox "There was an error"
End Sub



Private Sub Form_Close()
    On Error Resume Next
    If Me.MLOrig = "FA" Then Call CheckATFChanges(Me.SGrpID, "Economics Change")
    If Me.MLOrig = "FA" Then Call CheckATFChanges(Me.SGrpID, "Lease Terms Change")
FixMe:
End Sub



Private Sub ExitOptionsConsole_Click()
    If DLookup("MLOrig", "vw_FiveKeys", "SchID=" & Me.SchID & "") <> "FA" Then
        MsgBox "Only FA units can have EXIT Options", , "Not One of Ours"
        Exit Sub
    End If
    DoCmd.close acForm, "ExitOptionsManager"
    DoCmd.OpenForm "ExitOptionsManager"
    Forms!ExitOptionsManager.ClientList = Me.ClientID
    Form_ExitOptionsManager.ClientList_Click
    Forms!ExitOptionsManager.SchGrpList = Me.GroupID
    Form_ExitOptionsManager.SchGrpList_Click
End Sub


Private Sub ExitOptionsHistory_Click()
    If DLookup("MLOrig", "vw_FiveKeys", "SchID=" & Me.SchID & "") <> "FA" Then
        MsgBox "Only FA units can have EXIT Options", , "Not One of Ours"
        Exit Sub
    End If
    SchIDz = Me.SchID
    DoCmd.close acForm, "EXITOptions"
    DoCmd.OpenForm "EXITOptions"
    Forms!ExitOptions.ClientList = DLookup("ClientId", "vw_SixKeys", "SchID=" & SchIDz & "")
    Form_EXITOptions.ClientList_Click
    Forms!ExitOptions.LeaseList = DLookup("MLID", "vw_SixKeys", "SchID=" & SchIDz & "")
    Form_EXITOptions.Leaselist_Click
    Forms!ExitOptions.ScheduleList = SchIDz
    Form_EXITOptions.ScheduleList_Click
    Forms!ExitOptions.GroupList = DLookup("SGrpId", "vw_SixKeys", "SchID=" & SchIDz & "")
    Form_EXITOptions.History_Click
End Sub

Private Sub FinalEconomicsReportBtn_Click()
    If Me.Form.Dirty = True Then Me.Form.Dirty = False
    Call FinalEconomicsReport(Me.SchID, Me.GroupID, Me.SGrpNO)
  End Sub


Public Sub checkPO()
    On Error GoTo checkPO_Error

    If Me.PONum1 & "test" = "test" Then
        Dim NextSeq As Integer
        NextSeq = CInt(DMax("PoSeq", "Last10PONumbers")) + 1
        Dim ResposeX As VbMsgBoxResultEx
        Dim msg1 As String
        msg1 = msg1 & "<font size=3>Asign a PO Number</font><br />The Next Valid PO Number is:"
        msg1 = msg1 & "<br/><b>" & Me.SGrpFAID & "-" & NextSeq & "</b>"
        msg1 = msg1 & "<br /> would you like to assign it to this group?<br />"
        msg1 = msg1 & "   Please select..."
        ResposeX = Dialog.RichBox(msg1, vbCustom + vbExclamation, AtlaasVersion, , , 0, 1, 0, 0, "Yes", "No")
        If ResposeX = vbBt2 Then Exit Sub
        If ResposeX = vbBt1 Then
            Me.PONum1 = Me.SGrpFAID & "-" & NextSeq
            If Me.Form.Dirty Then Me.Form.Dirty = False
        End If
    End If

    On Error GoTo 0
    Exit Sub

checkPO_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkPO of VBA Document Form_GroupForm"
End Sub

Private Sub MakePaiLetter_Click()
   Call MakePAI(Me.SchID)
End Sub

Private Sub MakePO_Click()
    On Error GoTo FixMe
    Call checkPO
    If Me.PONum1 & "test" = "test" Then
        MsgBox "Sorry but you haven't assigned a PO number" & vbNewLine & "The reports will not run", vbCritical, AtlaasVersion
        Exit Sub
    End If
    DoCmd.close acReport, "PO_Vendor"
    DoCmd.OpenReport "PO_Vendor", acViewPreview
    On Error GoTo 0
    Exit Sub
FixMe:
    MsgBox "'Unit PO' Was Canceled, " & vbNewLine & "Error " & Err.Number & " (" & Err.Description & ") in procedure MakePO_Click of VBA Document Form_GroupForm", , AtlaasVersion
End Sub

Private Sub MakePOTracker_Click()
    On Error GoTo FixMe
    Call checkPO
    If Me.PONum1 & "test" = "test" Then
        MsgBox "Sorry but you haven't assigned a PO number" & vbNewLine & "The reports will not run", vbCritical, AtlaasVersion
        Exit Sub
    End If
    DoCmd.close acReport, "PO_VendorTracking"
    DoCmd.OpenReport "PO_VendorTracking", acViewPreview
    On Error GoTo 0
    Exit Sub
FixMe:
    MsgBox "'PO Tracker' Was Canceled, " & vbNewLine & "Error " & Err.Number & " (" & Err.Description & ") in procedure MakePOTracker_Click of VBA Document Form_GroupForm", , AtlaasVersion

End Sub

Private Sub OpnDrawDownManager_Click()
    DoCmd.OpenForm "DrawDownManager"
End Sub

Private Sub ReasonNote_DblClick(Cancel As Integer)
    If Me.ReasonINDX & "test" = "test" Then Exit Sub
    DoCmd.close acForm, "ATFReasonEdit"
    DoCmd.OpenForm "ATFReasonEdit", , , "INDx =" & Me.ReasonINDX
End Sub

Private Sub SgrpBnkNo_AfterUpdate()
    If Me.SgrpBnkNo Like "*-*" Then
        MsgBox "You cannot use a dash, just the number please.", vbCritical, "Wrong Fromat"
        Me.SgrpBnkNo = Replace(Me.SgrpBnkNo, "-", "")
    End If
End Sub

Private Sub SGrpFSLPmt_AfterUpdate()
    If Me.SGrpFSLPmt > 0 Then Me.PVFeesYN = True Else Me.PVFeesYN = True
End Sub

Private Sub SGrpNO_AfterUpdate()
    Committ = MsgBox("You're about to Change the Group No" & vbNewLine & _
    "" & vbNewLine & "From: " & PrevX & vbNewLine & _
    "" & vbNewLine & "To: " & Me.SGrpNO.Text & vbNewLine & vbNewLine & "Are you sure?", vbQuestion + vbYesNo + vbCritical, "Set CofA Signer")
    If Committ = vbNo Then
       Me.Undo
    End If
    If DCount("sgrpid", "SchGrp", "SGrpNO='" & Me.SGrpNO.Text & "' AND SGrpSchID= " & Me.SchID) > 0 Then
       MsgBox "'" & Me.SGrpNO.Text & "' is already in use in this Schedule." & vbNewLine & "Undoing the change" & _
                 "" & vbNewLine & vbNewLine & "Please try again", vbCritical, "You'll Hurt Yourself"
       Me.Undo
    End If
    PrevX = ""
    End Sub


Private Sub SGrpNO_GotFocus()
   PrevX = Me.SGrpNO
End Sub

Private Sub SyndCheckExcel_Click()
    Call SynCheckExcel(Me.SchID)
End Sub

Private Sub UnitRevDetail_DblClick(Cancel As Integer)
    If Me.UnitRevDetailLabel.caption Like "*rent*" Or Me.UnitRevDetailLabel.caption Like "*Unit*" Then
        Call OpenUnitsForm(Me.UnitRevDetail.Column(0))
        'DoCmd.OpenForm "UnitsForm", , , "UNITID=" & Me.UnitRevDetail.Column(0)
    End If
End Sub

Private Sub UseSpecialRemitGrp_Click()
    If UseSpecialRemitGrp = True Then Me.SpecialRemitAddrGrp.Locked = False Else Me.SpecialRemitAddrGrp.Locked = True
End Sub
Private Sub Validate_Click()
    Dim Msgx As String
    DoCmd.Save acForm, Me.Form.Name

    If DLookup("MLOrig", "vw_FiveKeys", "GroupID=" & Me.GroupID) = "FA" Then
        If Me.AssignedIndicator = "F" Then Msgx = "... then again, you already knew that :)"
        If GroupFinalizedCheck(Me.GroupID) = "Good" Then MsgBox "This Group checks out 'Good'" & vbNewLine & " (you may finalize)" & Msgx, vbInformation, AtlaasVersion
    Else
        MsgBox "This is not a 'FA' Group" & vbNewLine & "(you cannot finalized a tracked Units)", vbCritical, AtlaasVersion
        Me.FundDate.Undo
        Me.FundDate = Null
    End If
    Me.Form.Repaint
End Sub
Private Sub AssignedIndicator_GotFocus()
    CurrentIndicator = Me.AssignedIndicator
End Sub
Private Sub AssignedIndicator_Click()
    If Me.Form.NewRecord Then Exit Sub
    DoCmd.Save acForm, Me.Form.Name
    If Me.MLOrig = "FA" Then
        If Me.AssignedIndicator = "F" Then
            'And Nz(Me.BLCD, #1/1/2050#) > #1/25/2016# Then  'Ignore older Groups
            'And GroupSignOffCheck(Me.groupID, "SignOff6") Like "*has NOT detected*"
            If GroupFinalizedCheck(Me.GroupID) = "Good" Then
                Call MarkGroupDraftAsPaid(Me.GroupID)
                Call StartSignOff(Me.GroupID, "SignOff6")
                Call StartSignOff(Me.GroupID, "SignOff7")
            Else
                Me.AssignedIndicator = CurrentIndicator
            End If
        End If
        Call CheckFunded
    Else
        MsgBox "This is not a 'FA' Group" & vbNewLine & "(you cannot finalized a tracked Units)", vbCritical, AtlaasVersion
    End If
End Sub

Public Sub CheckFunded()
'** Final validation of Business rules
'(Changed 3/21/2018)
    On Error Resume Next
    Msgx = ""
    If AssignedIndicator = "F" Then
        Me.SyndActualLbl1.caption = "Synd Final"
        Me.SyndActualLbl2.caption = "Synd Final"
        Call LockUnLock(True)
        If IsGroupBLCDSet(Me.GroupID) = True And WasGroupBLCDNoticeSent(Me.GroupID) = False Then Call BLCDSyndFinalMessage(Me.GroupID)
    Else
        Me.SyndActualLbl1.caption = "Synd Projected"
        Me.SyndActualLbl2.caption = "Synd Projected"
        Call LockUnLock(False)
    End If
    Me.Assigned_Indicator = Me.AssignedIndicator
    Me.LeasePymtForGroup.SetFocus

    If FAUserAccessGroup Like "*SYND*" Then
        If Me.FundDate_Projected = Date And Me.Assigned_Indicator <> "F" Then Msgx = Msgx & "<b>Reminder.</b><Br/>" & _
           "The<b> Target Fund Date</b> is today and,<Br/>the<b> Actual Fund Date</b> has not been entered.<Br/><Br/>Enter the Actual Fund Date or Change the Target Fund Date."
        If Me.FundDate_Projected < Date And Me.Assigned_Indicator <> "F" Then Msgx = Msgx & "<b>Alert!</b><Br/>" & _
           "The<b> Target Fund Date</b> has Passed,<Br/>and the<b> Actual Fund Date</b> has not been entered.<Br/><Br/>Enter the Actual Fund Date or Change the Target Fund Date."
    End If
    If Msgx <> "" Then
        Me.SyndMsg = Msgx
        Call Dialog.RichBox(Msgx, vbCustom + vbCritical, "Target Fund Date issue!", , , 0, 1, 0, 0, "OK")

      '  MsgBox Msgx, vbCritical, "Error"
        Call LockUnLock(False)
    End If
End Sub

Public Sub LockUnLock(LockIt As Boolean)
'On Error Resume Next
    If LockIt = True Then
        Dim ctrl As Control
        For Each ctrl In Me.Controls
            If TypeOf ctrl Is TextBox Or ctrl.Tag = "Protected" Then ctrl.Locked = True
        Next
        Me.FormLocked.caption = "Form Locked"
        ValueX = "F"
        Me.LockUnLockBtn.Picture = "C:\FAAS\Images\Locked.png"
    Else
        For Each ctrl In Me.Controls
            If TypeOf ctrl Is TextBox Or ctrl.Tag = "Protected" Then ctrl.Locked = False
        Next
        Me.FormLocked.caption = "Form Un-Locked"
        Me.LockUnLockBtn.Picture = "C:\FAAS\Images\UnLocked.png"
        ValueX = "P"
    End If
    Me.PVRate.Locked = False
    Me.InterimLoanRate.Locked = False
    Me.LeasePymtForGroup.SetFocus

End Sub
Private Sub FundDate_BeforeUpdate(Cancel As Integer)
    If Me.FundDate <> "" Then
        If GroupFinalizedCheck(Me.GroupID) <> "Good" Then
            Call Dialog.RichBox("Sorry, but you can't enter a <b>Final Fund Date</b> just yet", vbCustom + vbCritical, "A Final Fund Date = Syndication!", , , 0, 1, 0, 0, "Undo")
            Me.Form.Undo
        Else
            Call Dialog.RichBox("This Group will be marked as <b>Finalized</b>,<br/> ", vbCustom + vbInformation, "Congrats! on the Syndication!", , , 0, 1, 0, 0, "Done")
            Me.Assigned_Indicator = "F"
            Call CheckFunded
        End If
    End If
End Sub
Private Sub FundDate_AfterUpdate()
'revised due to on-going Business Rule violations (2/20/2019) lbj
'   Call CheckFunded
'   Dim ResposeX As VbMsgBoxResultEx
'   Dim Msgx   As String
'   If Me.Assigned_Indicator <> "F" Then
'   If Me.FundDate < Date Then
'       Msgx = "The<b> Fund Date</b> you entered is a passed date,<Br/>and the Group has not been<b> Finalized</b>." & _
        '              "<Br/><Br/>Enter a valid Fund Date or leave it blank until funded."
'       ResposeX = Dialog.RichBox(Msgx, vbCustom + vbExclamation, AtlaasVersion, , , 0, 1, 0, 0, "Undo Date")
'       Me.Form.Undo
'   End If
'   Msgx = "Is this Group <b>Finalized</b> Yes or No?<br/>" & _
    '          " If '<b>No</b>', then the Fund Date will be removed.<br/>" & _
    '          " If '<b>Yes</b>' then this Group will be <b>F<b>inalized</b> and<b> Locked</b>."
'   ResponseX = Dialog.RichBox(Msgx, vbQuestion, "A Final Fund Date = Syndication!", , , 0, 1, 0, 0, "Yes", "No")'
'   If ResposeX = vbBt2 Then Me.FundDate = Null
'   If ResposeX = vbBt1 Then
'       Me.Assigned_Indicator = "F"
'       Call Validate_Click
'   End If
'  End If
'  Call CheckFunded
End Sub

Private Sub GoHome_Click()
    DoCmd.OpenForm "FAMainMenu"
End Sub

Private Sub LockUnLockBtn_Click()
    If Me.FormLocked.caption = "Form Locked" Then
        DoCmd.OpenForm "OPSPassword", , , , , , Me.formName
    Else
        Call LockUnLock(True)
    End If

End Sub


Private Sub ShowGroupSelected()
    Dim i As Integer
    For i = 0 To GroupsList.ListCount
        If (Me.GroupsList.Column(0, i)) = Nz(Me.OpenArgs, Me.UnitGroupID) Then
            Me.GroupsList.Selected(i) = True
            Me.GroupsList.SetFocus
            '  Me.GroupsList.ListIndex = (I)
        End If
    Next i
End Sub
Private Sub HideShowUAT()
    If FAUserAccessGroup Like "*UAT*" Then
        Dim Ccrl As Control
        For Each Ccrl In Me.Controls
            If Me(Ccrl.Name).Tag = "UAT" Then
                Me(Ccrl.Name).Visible = True
            End If
        Next Ccrl
    End If
End Sub



Private Sub PVFeesYN_Click()
    If Me.PVFeesYN = True Then Me.ServFeePassTroughYN = False
End Sub

Private Sub ReFreshMe_Click()
    Call CalcSyndFees
    Me.Form.Requery
End Sub

Private Sub ServFeePassTroughYN_Click()
    If Me.ServFeePassTroughYN = True Then Me.PVFeesYN = False
End Sub

Private Sub IndexRate_AfterUpdate()
    Call CalcSyndFees
End Sub

Private Sub SgrpSalesYield_AfterUpdate()
    Call CalcSyndFees
End Sub

Private Sub SgrpYield_AfterUpdate()
    Call CalcSyndFees
End Sub

Private Sub SalePrice_AfterUpdate()
    Call CalcSyndFees
End Sub

Private Sub Fee_Projected_AfterUpdate()
    Call CalcSyndFees
End Sub


Private Sub SgrpSynPenalty_AfterUpdate()
    Call CalcSyndFees
End Sub

Private Sub HoldbackConversion_AfterUpdate()
    Call CalcSyndFees
End Sub

Private Sub HoldbackResidual_AfterUpdate()
    Call CalcSyndFees
End Sub

Private Sub PerDiemProjected_AfterUpdate()
    Call CalcSalesFees
    Call CalcSyndFees
End Sub

Private Sub SgrpTermPenalty_AfterUpdate()
    Call CalcSalesFees
End Sub

Private Sub PerdiemActual_AfterUpdate()
    Call CheckPerdiem
    If Abs(Me.PerDiemActual - Round(Me.PerDiemCalculated, 2)) <> 0 And Me.PerDiemCalculated > 0 Then
        MsgBox "You're entering a Per Diem Per Unit Value: " & Format(Me.PerDiemActual, "$#,#0.#0") & vbNewLine & _
               "Which differes from the calculated amount of: " & Format(Me.PerDiemCalculated, "$#,#0.#0") & vbNewLine & vbNewLine, vbInformation, "Per Diem Not Actual"
    End If
    Call CalcSyndFees
End Sub
Private Sub CheckPerdiem()
    Me.PerDiemCalculated = CDbl(Nz(DLookup("PerDiemUnit", "PerDiem_Days_Rate_Group", "GroupRef=" & Me.GroupID), 0))
    If Nz(Me.PerDiemActual, 0) <> Me.PerDiemCalculated And Me.PerDiemCalculated > 0 Then
        Me.PerDiemActual.BackColor = RGB(255, 255, 153)
    Else
        Me.PerDiemActual.BackColor = RGB(255, 255, 255)
    End If
End Sub

Private Sub WithHeldPymts_AfterUpdate()
    Call CalcSyndFees
End Sub

Private Sub FundDate_Projected_AfterUpdate()
    If Me.Form.NewRecord Then Exit Sub
    Call CheckFunded
End Sub

Private Sub SyndMiscLabel1_AfterUpdate()
    If Me.SyndMiscLabel1 & "test" = "test" And Me.SyndMiscValue1 <> 0 Then
        MsgBox "You must have a label if you have a Misc Value 1 <> 0", , "You really need a label."
        Me.SyndMiscLabel1 = "Misc Value 1"
    End If
End Sub
Private Sub SyndMiscLabel2_AfterUpdate()
    If Me.SyndMiscLabel2 & "test" = "test" And Me.SyndMiscValue2 <> 0 Then
        MsgBox "You must have a label if you have a Misc Value 2 <> 0", , "You really need a label."
        Me.SyndMiscLabel1 = "Misc Value 2"
    End If
End Sub

Private Sub SyndMiscValue1_AfterUpdate()
    Call CalcSyndFees
    If Me.SyndMiscLabel1 & "test" = "test" Then
        MsgBox "Please type in a label for this 'misc' value", , "Need a label please"
        Me.SyndMiscLabel1 = "Misc Value 1"
    End If
End Sub
Private Sub SyndMiscValue2_AfterUpdate()
    Call CalcSyndFees
    If Me.SyndMiscLabel2 & "test" = "test" Then
        MsgBox "Please type in a label for this 'misc' value", , "Need a label please"
        Me.SyndMiscLabel2 = "Misc Value 2"
    End If
End Sub

Private Sub SynFee_AfterUpdate()
    Me.SalePrice = [CalcOEC] + [SynFee]
End Sub




Private Sub CalcSyndFees()
    If DLookup("MLOrig", "vw_FiveKeys", "GroupID=" & Me.GroupID) = "FA" Then Me.SalePrice = [CalcOEC] + [SynFee]
    Me.SynFee.Requery
    Me.SyndNetFee.Requery
    Me.ActualServiceFee.Requery
    Me.WithheldPayments.Requery
    Me.NetBankPayment.Requery
    Me.StripsNum.Requery
    Me.spread.Requery
    Me.SpreadIRR.Requery
    Me.StripsAmount.Requery
End Sub
Private Sub CalcSalesFees()
    Me.ProjectedNetFee.Requery
    Me.SalesSpread.Requery
End Sub
Private Sub ResidualGuaranteeDocYN_Click()
    Call CheckDocsFields(Me.ActiveControl.Name, "RGD")
End Sub

Private Sub ShortFallAgreementYN_Click()
    Call CheckDocsFields(Me.ActiveControl.Name, "SFA")
End Sub
Private Sub IndemnificationDocYN_Click()
    Call CheckDocsFields(Me.ActiveControl.Name, "IDF")
End Sub
Private Sub CustomDocumentNameYN_Click()
    Call CheckDocsFields(Me.ActiveControl.Name, "CustDoc")
End Sub
Public Sub CheckDocsFields(ParentControlX As String, TagX As String, Optional ValdateCheckBoxes As Boolean)
    Dim Ccrl As Control
    TickedX = "No"
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = TagX Then
            If Me(ParentControlX) = True Then Me(Ccrl.Name).Enabled = True Else Me(Ccrl.Name).Enabled = False
            If ValdateCheckBoxes = True Then
                If Me(Ccrl.Name).Value = True Then TickedX = "Yes"
            End If
        End If
    Next Ccrl
    If Me(ParentControlX) = True And ValdateCheckBoxes = True And TickedX = "No" Then
        MsgBox "Warning! Special Assignment Agreement:" & vbNewLine & ParentControlX & " is selected " & vbNewLine & "But, none of the sub-items are 'checked'." & vbNewLine & "Please correct.", vbExclamation, "No Items ticked"
        TickedX = "Stop"
        Me(ParentControlX).SetFocus
    End If
End Sub

Private Sub BLCDHelp_Click()
    Dim Msgx As String
    Msgx = "Negotiated = The targeted BLCD was negotiated w/the customer." & vbNewLine & "The actual BLCD was set accordingly as negotiated." & vbNewLine & vbNewLine
    Msgx = Msgx & "Per Lease = The targeted BLCD was set in accordance with the Lease." & vbNewLine & "The actual BLCD was set in accordance with the Lease." & vbNewLine & vbNewLine
    Msgx = Msgx & "Per ExchangeIT = The targeted BLCD was set in accordance w/ExchangeIT because it is a Replacement Schedule for an 'ExchangeIT' of an earlier Schedule" & vbNewLine & _
           "and such terms of that ExchangeIT will dictate how the BLCD will be set, which will be in accordance with the last rent payment of the schedule it is replacing." & vbNewLine & "The actual BLCD was set accordingly." & vbNewLine
    Msgx = Msgx & "   Example:" & vbNewLine & "   If last rent payment date is 10/25/15 for rental period of: 10/25/15 thru 11/24/15," & vbNewLine & "   then the per diem start date of the new schedule is usually 11/25/15 and the" & vbNewLine & "   BLCD is the 25th of the month following, which would be 12/25/15."
    MsgBox Msgx, , "'BLCD Set-By' Help"
End Sub



Private Sub MTMRate_AfterUpdate()
    MsgBox "This is a Calculated Field..." & vbNewLine & "'M-T-M Rent' is set at the unit level", vbCritical, "Use 'Mass-Mod'"
End Sub
Private Sub MTMRent_AfterUpdate()
    MsgBox "This is a Calculated Field..." & vbNewLine & "'M-T-M Rent' is set at the unit level", vbCritical, "Use 'Mass-Mod'"
End Sub
Private Sub LeaseExtOption1LRF_AfterUpdate()
    MsgBox "This is a Calculated Field..." & vbNewLine & "'Extended Rent' is set at the unit level", vbCritical, "Use 'Mass-Mod'"
End Sub

Private Sub LeaseExtOption1Rent_AfterUpdate()
    MsgBox "This is a Calculated Field..." & vbNewLine & "'Extended Rent' is set at the unit level", vbCritical, "Use 'Mass-Mod'"
End Sub
Private Sub SetGroupList()
    SQLz = "SELECT SGrpID, UnitGroup, Count(AssetID) AS Units " & _
           "FROM vw_SixKeys " & _
           "WHERE SchID=" & Me.SchID & " " & _
           "GROUP BY SGrpID, UnitGroup ORDER BY UnitGroup;"
    Me.GroupsList.RowSource = SQLz
    '
End Sub
Private Sub GroupsList_Click()
    On Error GoTo FixMe
    Me.Form.RecordSource = "Select * From GroupsFormQuery Where [SGrpID] =" & Me.GroupsList.Column(0) & ";"
    If Nz(Me.Assigned_Indicator, "P") = "F" Then Call LockUnLock(True) Else Call LockUnLock(False)
FixMe:
End Sub

Private Sub CloseBtn_Click()

End Sub



Private Sub AddIn_Fees_Click()
    If Me.ShowFee = "Hide" Then
        Me.SGrpFSLPmt.Visible = False
        Me.SGrpExtFSLPmt.Visible = False
        Me.SGrpRSLTPmt.Visible = False
        Me.ExtendFSLPymt.Visible = False
        Me.ServFeeLabel.Visible = False
        Me.AncillaryFeeLabel.Visible = False
        Me.EFSLLBL.Visible = False
        Me.TslTxLBL.Visible = False
        Me.ShowFee = "Show"
    Else
        Me.SGrpFSLPmt.Visible = True
        Me.SGrpExtFSLPmt.Visible = True
        Me.SGrpRSLTPmt.Visible = True
        Me.ExtendFSLPymt.Visible = True
        Me.ServFeeLabel.Visible = True
        Me.AncillaryFeeLabel.Visible = True
        Me.EFSLLBL.Visible = True
        Me.TslTxLBL.Visible = True
        Me.ShowFee = "Hide"
    End If

End Sub


Private Sub gCPNBtn_Click()

    DoCmd.close acForm, "GroupForm"
    DoCmd.close acForm, "Schedule_Form"
    DoCmd.close acForm, "schedules_sub"
    DoCmd.OpenForm "ClientsView", acNormal, "", "", , acNormal

End Sub

Private Sub FeeTerm_AfterUpdate()
    If Me.FeeTerm > LeaseTermPrimary Then
        MsgBox "You cannot enter a value greater than " & Me.LeaseTermPrimary & " months" & vbNewLine & "(the term of the lease).", vbCritical, "Do The Math"
        Me.FeeTerm.Value = 0
    End If

End Sub
Private Sub OpenExitOptions_Click()
    If DLookup("MLOrig", "vw_FiveKeys", "SchID=" & Me.SchID & "") <> "FA" Then
        MsgBox "Only FA units can have EXIT Options", , "Not One of Ours"
        Exit Sub
    End If
    SchIDz = Me.SchID
    DoCmd.close acForm, "EXITOptions"
    DoCmd.OpenForm "EXITOptions"
    Forms!ExitOptions.ClientList = DLookup("ClientId", "vw_SixKeys", "SchID=" & SchIDz & "")
    Form_EXITOptions.ClientList_Click
    Forms!ExitOptions.LeaseList = DLookup("MLID", "vw_SixKeys", "SchID=" & SchIDz & "")
    Form_EXITOptions.Leaselist_Click
    Forms!ExitOptions.ScheduleList = SchIDz
    Form_EXITOptions.ScheduleList_Click
    Forms!ExitOptions.GroupList = DLookup("SGrpId", "vw_SixKeys", "SchID=" & SchIDz & "")
    Form_EXITOptions.GroupList_Click
End Sub

Private Sub OpenMassMod_Click()
    DoCmd.close acForm, "BulkChange_form"
    DoCmd.OpenForm "BulkChange_form"
    Forms!BulkChange_Form.ClientList = Me.ClientID
    Form_BulkChange_Form.ClientList_Click
    Forms!BulkChange_Form.LeaseList = DLookup("MLID", "vw_SixKeys", "SchId=" & Me.SchID & "")
    Form_BulkChange_Form.Leaselist_Click
    Forms!BulkChange_Form.ScheduleList = Me.SchID
    Form_BulkChange_Form.ScheduleList_Click
    Forms!BulkChange_Form.GroupList = Me.GroupID
    Form_BulkChange_Form.GroupList_Click
    Form_BulkChange_Form.GetAll_Click
End Sub

Private Sub Assigned_Indicator_DblClick(Cancel As Integer)
    DoCmd.OpenForm "Unlock Grp"
    Me.SchNo.SetFocus
End Sub


Private Sub SelReturns_Click()

    DoCmd.OpenQuery "SetSUFSentq", acViewNormal, acEdit
    DoCmd.OpenQuery "RevPrintSUF", acViewNormal, acEdit

End Sub

Private Sub SUF_Selection_Click()
    DoCmd.OpenForm "SUF_Select_Form"
    Forms!SUF_Select_Form.Company_List = Me.ClientID
    Form_SUF_Select_Form.Company_List_Click
End Sub

Private Sub ThirdPartyPayAssgnmtYN_Click()
    If Me.ThirdPartyPayAssgnmtYN = True Then Me.FA3rdPartyPayingAgent = "Yes" Else Me.FA3rdPartyPayingAgent = "No"
End Sub
Private Sub FA3rdPartyPayingAgent_Click()
    If Me.FA3rdPartyPayingAgent = "Yes" Then Me.ThirdPartyPayAssgnmtYN = True Else Me.ThirdPartyPayAssgnmtYN = False
End Sub

Private Sub UnitsInGroup_AfterUpdate()
    Me.Form.Refresh
    Me.GroupsList.Requery
End Sub

Private Sub SyndGroupCheck_Click()
    On Error GoTo SchChkRBtn_Click_Error
    DoCmd.close acReport, "GroupReviewSynd"
    DoCmd.OpenReport "GroupReviewSynd", acViewPreview, "", "", acNormal
    On Error GoTo 0
    Exit Sub
SchChkRBtn_Click_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SchChkRBtn_Click of VBA Document Form_GroupForm", , "ATLAAS Core"
End Sub

Private Sub SelectUpdateUnits_Click()
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "Rmkt Updt", acViewNormal, acEdit
    DoCmd.MoveSize 0, 0
    DoCmd.SetWarnings True
End Sub



Private Sub PCofABTN_Click()

    If (Eval("[Forms]![groupform]![mlno] Is Null Or [Forms]![groupform]![mldt] Is Null")) Then
        Beep
        MsgBox "The CofA requires the Master Lease No and Date be entered on the Mastser Lease Form", vbCritical, "Missing Master Lease Info"
        Exit Sub
    End If
    If (Forms!GroupForm!SGRPPLB = False And Forms!GroupForm!ClientShNm = "Wal-Mart Trans") Then
        DoCmd.OpenReport "CofA-WalMart", acViewPreview, "", "", acNormal
        Exit Sub
    End If
    If (Forms!GroupForm!SGRPPLB = False And Forms!GroupForm![client group] = "maines") Then
        DoCmd.OpenReport "CofA-Maines", acViewPreview, "", "", acNormal
        Exit Sub
    End If
    If (Forms!GroupForm!SGRPPLB = False And Forms!GroupForm![client group] = "Toyota Motor Sales") Then
        Beep
        MsgBox "CofA requires Schedule date and negotiated BLCD on the LO Economics tab to be entered", vbExclamation, "Toyota Required Info for CofA"
        DoCmd.OpenReport "CofA-Toyota", acViewPreview, "", "", acNormal
        Exit Sub
    End If
    If (Forms!GroupForm!SGRPPLB = True) Then
        'DoCmd.OpenReport "CofA PLB", acViewPreview, "", "", acNormal
        'Exit Sub
        DoCmd.OpenReport "CofA", acViewPreview, "", "", acNormal
        ' BP West Coast Products LLC
        Exit Sub
    End If
    If (Forms!GroupForm!SGRPPLB = False And Forms!GroupForm!ClientShNm = "BP PRODUCTS NA") Then
        DoCmd.OpenReport "CofaBPWest", acViewPreview, "", "", acNormal
        DoCmd.OpenReport "CofABPWestAttch", acViewPreview, "", "", acNormal
        Exit Sub
    End If
    DoCmd.OpenReport "CofA", acViewPreview, "", "", acNormal
    Exit Sub

End Sub
