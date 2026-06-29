Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz       As String
Dim RST        As DAO.Recordset
Dim stage      As Integer

'Private Sub ATFAdmin_Click()
'    Me.ATFAdmin3 = Me.ATFAdmin.Column(1)
'End Sub
'
'Private Sub ATFDate_AfterUpdate()
'    Me.ATFDate3 = Me.ATFDate
'End Sub

'Private Sub BDEX_Click()
'    Me.BDEX3 = Me.BDEX2
'End Sub

Private Sub CoLesseeID_Click()
    Me.CoLesseeID3 = Me.CoLesseeID.Column(1)
End Sub
Private Sub SubLessee1_AfterUpdate()
    Me.SubLessee = Me.SubLessee1
End Sub

Public Sub FAIDLPAFList_Click()
    Me.NumOfSchedules = FAIDLPAFList.Column(3)
    Me.NumOfGroups = FAIDLPAFList.Column(4)
    Me.NumOfUnits = FAIDLPAFList.Column(5)
    SQLz = SplitString(Me.Form.RecordSource, "Where", 0) & "Where vw_SFKeys.OpportunityID='" & FAIDLPAFList.Column(7) & "'"
    Me.Form.RecordSource = SQLz
    Me.whitebox.Visible = False
    Me.SF_ClientLegalName.Visible = True
    '   Set RST = Me.RecordsetClone
    '   RST.FindFirst "FAID = '" & Me.FAIDLPAFList.Column(0) & "'"
    '   Me.Bookmark = RST.Bookmark
    If Me.AccountID <> Me.SFAccountId Then
        'If Me.SF_ClientLegalName <> Me.ATLAASClient Then
        Dim Msgx, ResposeX As String
        Msgx = "The SF Client and the ATLAAS Core Client do <b>not</b> match<br /> Proceed with import?"
        ResposeX = Dialog.RichBox(Msgx, vbCustom + vbExclamation, AtlaasVersion, , , 0, 1, 0, 0, "Yes", "No")
        If ResposeX = 102 Then
            Me.GoNext.Enabled = False
            Exit Sub
        Else
            Me.GoNext.Enabled = True
        End If
    Else
        Me.GoNext.Enabled = True
    End If
    Call SetButtons("GotoStep2")
    stage = 1

End Sub

Private Sub Form_Current()
    Me.MLID.Requery
    Me.MLID2.Requery
    Me.FAID.Requery
    Me.CoLesseeID.Requery
    Call Refresh_My_Buttons
    Call MakeMLIDArray
    Me.ATFAdmin = FAUserIDx
    Me.ATFAdmin3 = FAUserIDx
    Me.ATFDate = Date
    Me.ATFDate3 = Date

    If FAUserAccess < 99 Then
        Me.NumOfSchedules.Visible = False
        Me.NumOfGroups.Visible = False
        Me.SchedulesLabel.Visible = False
        Me.GroupsLabel.Visible = False
    Else
        Me.NumOfSchedules.Visible = True
        Me.NumOfGroups.Visible = True
        Me.SchedulesLabel.Visible = True
        Me.GroupsLabel.Visible = True
    End If
End Sub
Public Sub MakeMLIDArray()

    Me.MLID.RowSource = ""
    Dim MLSource As String
    Dim xi     As Integer
    SQLz = "SELECT MLID, MLNo, Client, MLOrig " & _
           "From MstrLease WHERE Client= " & Nz(Me.LesseeID, 0) & " AND  MLOrig ='FA' " & _
           "ORDER BY MstrLease.MLID DESC , MstrLease.MLNo;"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)


    MLSource = "0;'Create A Master Lease On Import';0"
    On Error GoTo FixMe
    RST.MoveFirst
    For xi = 1 To RST.RecordCount + 1
        MLSource = MLSource & ";" & RST!MLID & ";'" & RST!MLNo & "';" & RST!Client
        RST.MoveNext
    Next xi
FixMe:
    On Error GoTo 0
    Me.MLID.RowSource = MLSource
    RST.close
    Set RST = Nothing
End Sub
Public Sub AddMasterLease()
    Dim MLNox, LeaseExist As String
    MLNox = "MasterLease-" & Me.FAID
    If DCount("MLID", "MstrLease", "MLNo='" & MLNox & "' AND Client=" & Me.LesseeID) >= 1 Then
        MsgBox "There is already a Master Lease [ " & MLNox & " ] " & vbNewLine & "for " & Me.ClientSN & vbNewLine & vbNewLine & _
               "No Lease was added", vbCritical, atlassversion
        Refresh_My_Buttons
        LeaseExist = "Yes"
        Me.TabCtrl1.Pages.Item(2).SetFocus
        Exit Sub
    End If
    LeaseExist = "No"
    SQLz = " Insert into MstrLease (MLNo, MLOrig, MLDt, Client, MLSigner) Values ('" & MLNox & "', 'FA', #" & Date & "#, " & Me.LesseeID & ", '');"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Call MakeMLIDArray
    Me.MLID.Requery
    Me.MLID = DLookup("MLID", "MstrLease", "MLNo='" & MLNox & "' AND Client=" & Me.LesseeID)
    Me.MLID2 = Me.MLID.Column(1)
    Me.MLID3 = Me.MLID.Column(1)
    Refresh_My_Buttons
End Sub

Private Sub Form_Load()
    stage = 0
    Call ShowStep(1)
    Call SetButtons("All")
    Me.whitebox.Visible = True
    Me.SF_ClientLegalName.Visible = False
End Sub

Public Sub GoBack_Click()
    Select Case stage
    Case Is < 2
        Me.FAIDLPAFList = ""
        stage = 1
        Call ShowStep(1)
        Me.ImportToAtlaas.Visible = False
        Me.GoNext.Visible = True
    Case 2
        Me.TabCtrl1.Pages.Item(0).SetFocus
        stage = 1
        Call ShowStep(1)
        Me.MLID = ""
        Me.MLID2 = ""
        Me.MLID3 = ""
        Me.ImportToAtlaas.Visible = False
        Me.GoNext.Visible = True
    Case 3
        Me.TabCtrl1.Pages.Item(1).SetFocus
        stage = 2
        Call ShowStep(2)
        Me.CoLesseeID = Null
        Me.CoLesseeID3 = Null
        Me.SubLessee = Null
        Me.SubLessee1 = Null
        Me.GuarantorID = Null
        Me.GuarantorID3 = Null
        ' Me.ATFDate = ""
        ' Me.ATFDate3 = ""
        ' Me.ATFAdmin = ""
        ' Me.ATFAdmin3 = ""
        Me.ImportToAtlaas.Visible = False
        Me.GoNext.Visible = True
    Case 4
        Me.TabCtrl1.Pages.Item(2).SetFocus
        stage = 3
        Call ShowStep(3)
        Me.ImportToAtlaas.Visible = False
        Me.GoNext.Visible = True
    End Select


End Sub
Private Sub ShowStep(StepX As Integer)
    Me.StepLabel1.FontBold = False
    Me.StepLabel2.FontBold = False
    Me.StepLabel3.FontBold = False
    Me.StepLabel4.FontBold = False
    Me.StepLabel1.Enabled = False
    Me.StepLabel2.Enabled = False
    Me.StepLabel3.Enabled = False
    Me.StepLabel4.Enabled = False
    Me(CStr("StepLabel" & StepX)).Enabled = True
    Me(CStr("StepLabel" & StepX)).FontBold = True
End Sub
Private Sub SetButtons(Excludeme As String)
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Buttons" And Me(Ccrl.Name).Name <> Excludeme Then
            Me(Ccrl.Name).Enabled = False
            Me(Ccrl.Name).BackColor = RGB(236, 236, 236)
        End If
        If Me(Ccrl.Name).Tag = "Buttons" And Me(Ccrl.Name).Name = Excludeme Then
            Me(Ccrl.Name).Enabled = True
            Me(Ccrl.Name).BackColor = RGB(173, 192, 217)
        End If
    Next Ccrl

End Sub

Public Sub GoNext_Click()
    Select Case stage
    Case Is < 2
        If Nz(Me.FAIDLPAFList.Column(0), 0) = 0 Then
            MsgBox "You must first select a FAID.", vbCritical, AtlaasVersion
            Exit Sub
        End If
        Set RST = Me.RecordsetClone
        RST.FindFirst "FAID = '" & Me.FAIDLPAFList.Column(0) & "'"
        Me.Bookmark = RST.Bookmark
        RST.close
        Set RST = Nothing
        Me.TabCtrl1.Pages.Item(1).SetFocus
        Me.MLID.SetFocus
        Call MakeMLIDArray
        stage = 2
        Call ShowStep(2)
    Case 2
        If Me.MLID & "test" = "test" Then
            MsgBox "Before you can proceed, you must: " & vbNewLine & _
                   "1 - Select an existing Master Lease" & vbNewLine & _
                   "   or" & vbNewLine & _
                   "2 - Choose 'Create a Master Lease On Import'", vbCritical, AtlaasVersion
            Exit Sub
        End If
        Me.TabCtrl1.Pages.Item(2).SetFocus
        stage = 3
        Call ShowStep(3)
    Case 3
        Me.BDEX3 = Me.BDEX2
        Dim Msgx As String
         SQLz = "SELECT TOP 1 CDbl(NumInString(Left(Schedule,3))) As SchInt "
         SQLz = SQLz & "From vw_SixKeys WHERE MLOrig = 'FA' AND ClientId=" & Me.LesseeID & " "
         SQLz = SQLz & "GROUP BY CDbl(NumInString(Left(Schedule,3))) "
         SQLz = SQLz & "ORDER BY CDbl(NumInString(Left(Schedule,3))) DESC;"
         Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
         If RST.RecordCount < 1 Then
         Me.NewScheduleName = "1"
         Else
         Me.NewScheduleName = RST!schint + 1
         End If


        Me.TabCtrl1.Pages.Item(3).SetFocus
        stage = 4
        Call ShowStep(4)
        Me.ImportToAtlaas.Visible = True
        Me.GoNext.Visible = False
        Call SetButtons("ImportToAtlaas")
        'Me.NumOfUnits.SetFocus
    Case 4
    End Select
End Sub

Private Sub GuarantorID_Click()
    Me.GuarantorID3 = Me.GuarantorID.Column(1)
End Sub

Private Sub LPAFDate_AfterUpdate()
    Me.LPAFcomp3 = Me.LPAFDate
End Sub

Private Sub MLID_Click()
    Me.MLID2 = Me.MLID.Column(1)
    Me.MLID3 = Me.MLID.Column(1)
End Sub

Private Sub CheckFAID()
    If DCount("FAID", "Deals", "FAID='" & Me.FAID & "'") > 0 Then
        Exit Sub
    Else
        SQLz = "INSERT INTO deals ( FAID, transTypeRef, dealShortDesc, dealRFPDate, dealLPAFDue, dealLPAFcomp, dealUnitQty, clientRef, dealOutcome"
        SQLz = SQLz & ", dealOutcomeDate, dealEnteredDate, dealEstDelDate, dealMPY, dealTerm, dealUnitPrice, dealGrandTotal, DealLOEx, SFDCType ) "
        SQLz = SQLz & "SELECT Deal.FAID, 2 AS TransType, Deal.SchShortDesc, Deal.OutcomeDate AS RPF, Deal.LPAFDue, Deal.LPAF, Deal.UnitCount"
        SQLz = SQLz & ", Deal.clientID, 'Awarded' AS Outcome, Deal.OutcomeDate, Date() AS AddDate, Deal.EstDelv, gSpec.MPY, gSpec.Term, Deal.AvgUnitCost"
        SQLz = SQLz & ", Deal.DealAmount, NZ(Deal.BDEX,'HS'), Deal.DealType "
        SQLz = SQLz & "FROM vw_SFI_SchDealSpec As Deal "
        SQLz = SQLz & "INNER JOIN (Select OpportunityId,Max(allowedMPY) As MPY,Max(leaseterm) As Term "
        SQLz = SQLz & "FROM vw_SFI_GroupSpec Group By OpportunityId) as gSpec ON Deal.OpportunityId = gSpec.OpportunityId "
        SQLz = SQLz & "Where FAID='" & Me.FAID & "';"
        '       Sqlz = "INSERT INTO deals ( FAID, transTypeRef, dealShortDesc, dealRFPDate, dealLPAFDue, dealLPAFcomp, dealUnitQty, clientRef, dealOutcome,dealOutcomeDate, dealEnteredDate, dealEstDelDate, dealMPY, dealTerm, dealUnitPrice, dealGrandTotal, DealLOEx ,SFDCType) " & _
                "SELECT FAID, 2, dealShortDesc, dealRFPDate, dealLPAFDue, dealLPAFcomp, dealUnitQty, clientRef, 'Awarded', dealOutcomeDate, dealEnteredDate, dealEstDelDate, dealMPY, dealTerm, dealUnitPrice, dealGrandTotal, DealLOEx, [Type]" & _
                "FROM ATF_SF_ImportDeal Where FAID='" & Me.FAID & "';"
        DoCmd.SetWarnings False
          DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
    End If
End Sub

Private Sub ImportToAtlaas_Click()
    If Me.NewScheduleName & "test" = "test" Then
        MsgBox "Sorry, You MUST supply a Schedule Number for This Import", vbCritical, "What are you doing silly rabbit?"
        Exit Sub
    End If

    If Me.MLID.Column(0) = 0 Then Call AddMasterLease
    Call CheckFAID
    Dim ATFIDx As Integer
    Dim i, ii, iii As Integer
    Dim Msgx As String
    SQLz = "SELECT * From vw_SFI_UnitSpec WHERE FAID='" & Me.FAID & "';"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    If RST.RecordCount < 1 Then
        Msgx = Msgx & "Invalid SF Import Record" & vbNewLine
        GoTo Errorlevel1
    End If
    RST.MoveFirst
    For i = 1 To RST.RecordCount
        If RST!ModelYear & "test" = "test" Then Msgx = Msgx & "SF Spec Year" & vbNewLine
        If RST!Make & "test" = "test" Then Msgx = Msgx & "SF Spec Make" & vbNewLine
        If RST!Model & "test" = "test" Then Msgx = Msgx & "SF Spec Model" & vbNewLine
        RST.MoveNext
    Next i
Errorlevel1:
    RST.close
    Set RST = Nothing
    If Me.MLID & "test" = "test" Then Msgx = Msgx & "    Master Lease" & vbNewLine
    If Msgx <> "" Then
        MsgBox "You're missing the following data:" & vbNewLine & Msgx & "Please correct..." & vbNewLine & "Deal Not Imported", vbCritical, AtlaasVersion
        Exit Sub
    End If
    Msgx = ""
Insert_ATF:
    MessageUser "Open", "Creating ATF", "Lease Structure"
    If DCount("ATFID", "ATF", "LPAFDate=#" & Me.LPAFDate & "# AND FAID='" & Me.FAID & "'") > 0 Then
        MessageUser "Open", "ATF Already Exist", "Moving On..."
        GoTo Insert_Schedule_Profile
    Else
        MessageUser "Open", "Creating ATF", "Lease Structure"
    End If
    SQLz = "INSERT INTO ATF "
    SQLz = SQLz & "(FAID "
    SQLz = SQLz & ",ATFDate"
    SQLz = SQLz & ",ATFDesc"
    SQLz = SQLz & ",LPAFDate"
    SQLz = SQLz & ",MLID"
    SQLz = SQLz & ",LeaseType"
    SQLz = SQLz & ",CreditRank"
    SQLz = SQLz & ",BDEX"
    SQLz = SQLz & ",Admin"
    SQLz = SQLz & ",LesseeID"
    SQLz = SQLz & ",CoLesseeID"
    SQLz = SQLz & ",NumOfSchedules"
    SQLz = SQLz & ",NumOfGroups"
    SQLz = SQLz & ",NumOfUnits"
    SQLz = SQLz & ",LastModified)"
    SQLz = SQLz & " Values ("
    SQLz = SQLz & "'" & Me.FAID & "'"
    SQLz = SQLz & ", #" & Me.ATFDate & "#"
    SQLz = SQLz & ",'" & Me.FAID & "-" & Me.ActualCloseDate & "'"
    SQLz = SQLz & ",#" & Me.ActualCloseDate & "#"
    SQLz = SQLz & "," & Me.MLID & ""
    SQLz = SQLz & ",'" & Me.LeaseType & "'"
    SQLz = SQLz & ",'1'"
    SQLz = SQLz & ", " & Me.BDEX & ""
    SQLz = SQLz & ", " & Me.ATFAdmin & ""
    SQLz = SQLz & ", " & Me.LesseeID & ""
    SQLz = SQLz & ", " & Nz(Me.CoLesseeID, 0) & ""
    SQLz = SQLz & "," & Me.NumOfSchedules & ""
    SQLz = SQLz & "," & Me.NumOfGroups & ""
    SQLz = SQLz & "," & Me.NumOfUnits & ""
    SQLz = SQLz & ", #" & Now() & "#)"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
Insert_Schedule_Profile:
GET_New_ATFID:
    ATFIDx = DLookup("ATFID", "ATF", "FAID='" & Me.FAID & "'")
    If DCount("ATFID", "ATFScheduleSpec", "ATFID=" & ATFIDx) >= CInt(Me.NumOfSchedules) Then
        MessageUser "Update", "Schedule Profile Already Exist", "Moving On..."
        GoTo Insert_Group_Profile
    Else
        MessageUser "Update", "Creating ATF - Schedule", "Lease Structure"
    End If

    SQLz = "INSERT INTO ATFScheduleSpec ( ATFSchName"
    SQLz = SQLz & ", SchFAID"
    SQLz = SQLz & ", ATFID"
    SQLz = SQLz & ", MLID"
    SQLz = SQLz & ", FSLType"
    SQLz = SQLz & ", ATFSchSpecDesc"
    SQLz = SQLz & ", SchDescription"
    SQLz = SQLz & ", NumGroupsInSch"
    SQLz = SQLz & ", SchSpecInserted )"
    SQLz = SQLz & " SELECT [Forms]![ATF_SFImporter]![NewScheduleName] AS ATFSchName"
    SQLz = SQLz & ", FAID AS SchFAID"
    SQLz = SQLz & ", '" & ATFIDx & "' AS ATFID"
    SQLz = SQLz & ", '" & Me.MLID & "' AS MLID"
    SQLz = SQLz & ", 'Yes' AS FSLType"
    SQLz = SQLz & ", SchShortDesc AS ATFSchSpecDesc"
    SQLz = SQLz & ", '" & Me.FAID & " - Sch: " & Me.NewScheduleName & ""
    SQLz = SQLz & ", containing " & [SchShortDesc] & ")' AS SchDescription"
    SQLz = SQLz & ", SpecGroups AS NumGroupsInSch"
    SQLz = SQLz & ", '" & Now() & "' AS SchSpecInserted "
    SQLz = SQLz & "FROM vw_SFI_SchDealSpec "
    SQLz = SQLz & "WHERE FAID='" & Me.FAID & "' ;"


    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True

Insert_Group_Profile:
    If DCount("ATFID", "ATFGroupSpec", "ATFID=" & ATFIDx) >= CInt(Me.NumOfGroups) Then
        MessageUser "Update", "Group Profile Already Exist", "Moving On..."
        GoTo UpdateGroupNum    ' was Insert_Unit_Spec
    Else
        MessageUser "Update", "Creating ATF - Group Profile", "Just a Sec"
    End If

    SQLz = "INSERT INTO ATFGroupSpec (ATFID "
    SQLz = SQLz & ", LPAFID"
    SQLz = SQLz & ", ATFSchSpecID"
    SQLz = SQLz & ", ATFGroupSpecDesc"
    SQLz = SQLz & ", UnitSpecsInGroup"
    SQLz = SQLz & ", UnitsInGroup"
    SQLz = SQLz & ", GrpFAID"
    SQLz = SQLz & ", ExchangeITGroup"
    SQLz = SQLz & ", ExchangeITType"
    SQLz = SQLz & ", ExchangeITMonth"
    SQLz = SQLz & ", ExtendITOptionTerm"
    SQLz = SQLz & ", ExtendedRentAvg"
    SQLz = SQLz & ", PrimaryLeaseTerm"
    SQLz = SQLz & ", PrimaryLeaseNotice"
    SQLz = SQLz & ", LeasePymtFreq"
    SQLz = SQLz & ", LeasePaidAdvArr"
    SQLz = SQLz & ", SwapIndex"    ' Swap
    SQLz = SQLz & ", SwapDate"
    SQLz = SQLz & ", SwapRate"
    SQLz = SQLz & ", SwapTermYears"
    SQLz = SQLz & ", SwapYield"
    SQLz = SQLz & ", SwapBonus"
    SQLz = SQLz & ", SwapAdjustment"
    SQLz = SQLz & ", Fee_Projected"
    SQLz = SQLz & ", UATPerUnit"
    SQLz = SQLz & ", PLB"
    SQLz = SQLz & ", LeaseType"
    SQLz = SQLz & ", BLCDBasisID"
    SQLz = SQLz & ", PerDiemBasisID"
    SQLz = SQLz & ", MPYAllowance"
    SQLz = SQLz & ", HPYAllowance"
    SQLz = SQLz & ", AllowanceMethod"
    SQLz = SQLz & ", CPMileThreshold"
    SQLz = SQLz & ", CPHourThreshold"
    SQLz = SQLz & ", FundDate_Projected"
    SQLz = SQLz & ") SELECT DISTINCT " & ATFIDx & ""
    SQLz = SQLz & ", LPAFID"
    SQLz = SQLz & ", (SELECT TOP 1 ATFSchSpecID From ATFScheduleSpec WHERE ATFID=" & ATFIDx & ")"
    SQLz = SQLz & ",'zz' As ATFGroupSpecDesc"
    ' SQLz = SQLz & ", '" & Chr(64 + i) & "' As ATFGroupSpecDesc"
    SQLz = SQLz & ", 1 AS UnitSpecsInGroup"
    SQLz = SQLz & ", UnitsInGroup"
    SQLz = SQLz & ", FAID"
    SQLz = SQLz & ", ExchangeITGroup"
    SQLz = SQLz & ",'Specified Mo' AS ExchangeITType"
    SQLz = SQLz & ", ExchangeITMonth"
    SQLz = SQLz & ", ExtdLeaseTerm"
    SQLz = SQLz & ", ExtdLeaseRent"
    SQLz = SQLz & ", LeaseTermPrimary"
    SQLz = SQLz & ", 3 AS LeaseNotice"
    SQLz = SQLz & ", 'M' AS LeasePymtFreq"
    SQLz = SQLz & ", PayAdvArr"
    SQLz = SQLz & ", ''"
    SQLz = SQLz & ", SwapDate"
    SQLz = SQLz & ", SwapRate"
    SQLz = SQLz & ", SwapYear"
    SQLz = SQLz & ", SwapYield"
    SQLz = SQLz & ", SwapBonus"
    SQLz = SQLz & ", (SwapAdjustUp * 0.01)"
    SQLz = SQLz & ", SynFeeProjected"
    SQLz = SQLz & ", 0 AS UATPerUnit"
    SQLz = SQLz & ", 0 AS PLB"
    SQLz = SQLz & ", LeaseType"
    SQLz = SQLz & ", 2 AS BLCDBasisID"
    SQLz = SQLz & ", 1 AS PerDiemBasisID"
    SQLz = SQLz & ", AllowedMPY AS MilesPerYear"
    SQLz = SQLz & ", AllowedHPY AS HoursPerYear"
    SQLz = SQLz & ", 'P' AS AllowanceMethod"
    SQLz = SQLz & ", [AllowedMPY]*1.1 AS CPMileThreshold"
    SQLz = SQLz & ", [AllowedHPY]*1.1 AS CPHourThreshold"
    SQLz = SQLz & ", FundDate_Projected"
    SQLz = SQLz & " FROM vw_SFI_GroupSpec"
    SQLz = SQLz & " WHERE vw_SFI_GroupSpec.FAID='" & Me.FAID & "';"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True

UpdateGroupNum:
    SQLz = "SELECT * From ATFGroupSpec WHERE GRPFAID='" & Me.FAID & "';"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    RST.MoveFirst

    For i = 1 To Me.NumOfGroups
        SQLz = "Update ATFGroupSpec Set ATFGroupSpecDesc = '" & Chr(64 + i) & "' WHERE ATFGroupSpecID=" & RST!ATFGroupSpecID
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
        RST.MoveNext
    Next i
Insert_Unit_Spec:
    MessageUser "Update", "Creating ATF - Unit Specs", "This will be quick..."

    SQLz = "INSERT INTO ATFUnitSpec ( "
    SQLz = SQLz & " ATFID"
    SQLz = SQLz & ", LPAFID"
    SQLz = SQLz & ", clientID"
    SQLz = SQLz & ", ClientGroupId"
    SQLz = SQLz & ", ATFGroupSpecID"
    SQLz = SQLz & ", ATFUnitSpecDesc"
    SQLz = SQLz & ", NumUnitsInSpec"
    SQLz = SQLz & ", UnitModelYear"
    SQLz = SQLz & ", UnitMake"
    SQLz = SQLz & ", UnitModel"
    SQLz = SQLz & ", UnitType"
    SQLz = SQLz & ", UnitSubType"
    SQLz = SQLz & ", UnitConfiguration"
    SQLz = SQLz & ", Axle"
    SQLz = SQLz & ", PowerType"
    SQLz = SQLz & ", EngineType"
    SQLz = SQLz & ", TransmissionType"
    SQLz = SQLz & ", GVW"
    SQLz = SQLz & ", VendorID"
    SQLz = SQLz & ", UnitCost"
    SQLz = SQLz & ", SchRent"
    SQLz = SQLz & ", FirstExtendedRent"
    SQLz = SQLz & ", SalesResidual"
    SQLz = SQLz & ", UTARebate"
    SQLz = SQLz & ", EstimatedDeliveryDate"
    SQLz = SQLz & ", DeliveryLocationID"
    SQLz = SQLz & ", GarageLocationID"
    SQLz = SQLz & ", RegTitleLocationID"
    SQLz = SQLz & ", GroupID )"
    SQLz = SQLz & " SELECT " & ATFIDx & " AS ATFID"
    SQLz = SQLz & ", US.LPAFID"
    SQLz = SQLz & ", US.clientID"
    SQLz = SQLz & ", US.ClientGroupId"
    SQLz = SQLz & ", AG.ATFGroupSpecID"
    SQLz = SQLz & ", 'SPEC-' & [FAID] AS ATFUnitSpecDesc"
    SQLz = SQLz & ", US.UnitsInSpec"
    SQLz = SQLz & ", US.ModelYear"
    SQLz = SQLz & ", US.make"
    SQLz = SQLz & ", US.Model"
    SQLz = SQLz & ", US.UnitType"
    SQLz = SQLz & ", US.SubType"
    SQLz = SQLz & ", [CabConfig] & [ReeferType] & [Length] AS Config"
    SQLz = SQLz & ", US.Axle"
    SQLz = SQLz & ", US.Powertype"
    SQLz = SQLz & ", US.Engine"
    SQLz = SQLz & ", US.Transmission"
    SQLz = SQLz & ", US.GVW"
    SQLz = SQLz & ", US.VendorID"
    SQLz = SQLz & ", US.UnitCost"
    SQLz = SQLz & ", Nz(US.RentSch,0)"
    SQLz = SQLz & ", US.RentExtd"
    SQLz = SQLz & ", Nz(US.RVSch,0)"
    SQLz = SQLz & ", US.UTA"
    SQLz = SQLz & ", US.EstDelv"
    SQLz = SQLz & ", 0 AS Expr4"
    SQLz = SQLz & ", 0 AS Expr5"
    SQLz = SQLz & ", 0 AS Expr6"
    SQLz = SQLz & ", 0 AS GroupID "
    SQLz = SQLz & "FROM ATFGroupSpec As AG INNER JOIN vw_SFI_UnitSpec As US ON AG.LPAFID = US.LPAFId "
    SQLz = SQLz & "WHERE US.FAID= '" & Me.FAID & "';"

    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True

Deploy_Schedule:
    Dim SchSpecIDx, GroupSpecIDx, NumOfSch As Integer
    Dim MLNox As String
    MLNox = DLookup("MLNO", "MstrLease", "MLID=" & Me.MLID)

InsertSchedule:

    SQLz = SQLz & ", SubLeaseeX Text ( 255 )"
    SQLz = SQLz & ", LastSchNo Long;"
    SQLz = "INSERT INTO Schedule ( ATFID"
    SQLz = SQLz & ", SchNo"
    SQLz = SQLz & ", SchMLRef"
    SQLz = SQLz & ", ATFSchSpecID"
    SQLz = SQLz & ", SchFAID"
    SQLz = SQLz & ", FSLType"
    SQLz = SQLz & ", SchShortDesc"
    SQLz = SQLz & ", ApprovalRequest"
    SQLz = SQLz & ", ApprovalReply"
    SQLz = SQLz & ", SchBillAdd"
    SQLz = SQLz & ", DateSchAdded"
    SQLz = SQLz & ", SubLeasee )"
    SQLz = SQLz & "SELECT DISTINCT ATFID"
    SQLz = SQLz & ", ATFSchName"
    SQLz = SQLz & ", MLID"
    SQLz = SQLz & ", ATFSchSpecID"
    SQLz = SQLz & ", SchFAID"
    SQLz = SQLz & ", 'No' AS FSLType"
    SQLz = SQLz & ", SchDescription"
    SQLz = SQLz & ", ApprovalRequest"
    SQLz = SQLz & ", ApprovalReply"
    SQLz = SQLz & ", ScheduleBillingID"
    SQLz = SQLz & ", '" & Date & "' AS DateSchAdded"
    SQLz = SQLz & ", '" & Me.SubLessee & "' AS SubLeasee"
    SQLz = SQLz & " FROM ATFScheduleSpec"
    SQLz = SQLz & " WHERE SchFAID='" & Me.FAID & "';"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True

DeployGroups:
    Dim ScIDx As Integer
    Dim GrpIdx As Integer
    SQLz = ""
    SQLz = "INSERT INTO schgrp ( ATFID"
    SQLz = SQLz & ", SGrpNO"                    '2
    SQLz = SQLz & ", SGrpSchID"
    SQLz = SQLz & ", ATFGroupSpecID"
    SQLz = SQLz & ", UnitsInGroup"              '5
    SQLz = SQLz & ", OrigEquipCost"
    SQLz = SQLz & ", ExchangeITGroup"
    SQLz = SQLz & ", ExchangeITType"
    SQLz = SQLz & ", ExchangeITMonth"
    SQLz = SQLz & ", LeaseExtOption1Term"       '10
    SQLz = SQLz & ", LeaseExtOption1Rent"
    SQLz = SQLz & ", MTMRent"
    SQLz = SQLz & ", LeaseTermPrimary"
    SQLz = SQLz & ", LeaseTermNotice"
    SQLz = SQLz & ", LeasePymtFreq"             '15
    SQLz = SQLz & ", LeasePaidAdvArr"
    SQLz = SQLz & ", LeasePymtForGroup"
    SQLz = SQLz & ", LRFGroup"
    SQLz = SQLz & ", SGrpSalesIndex"
    SQLz = SQLz & ", SGrpSalesIndDt"            '20
    SQLz = SQLz & ", SgrpSalesIndRt"
    SQLz = SQLz & ", SGprSalesIndYr"
    SQLz = SQLz & ", SgrpSalesYield"
    SQLz = SQLz & ", SGrpBonus"
    SQLz = SQLz & ", SGRPPLB"                   '25
    SQLz = SQLz & ", LeaseType"
    SQLz = SQLz & ", SgrpCommInd"
    SQLz = SQLz & ", PerDiemStartsOn"
    SQLz = SQLz & ", AllowedHoursPer"
    SQLz = SQLz & ", SGprAllow1"                '30
    SQLz = SQLz & ", SGrpov1"
    SQLz = SQLz & ", SGprAllow2"
    SQLz = SQLz & ", SGrpov2"
    SQLz = SQLz & ", SGprAllow4"
    SQLz = SQLz & ", SGrpov4"                   '35
    SQLz = SQLz & ", SGprAllow5"
    SQLz = SQLz & ", SGrpov5"
    SQLz = SQLz & ", SGrpFSLPmt"
    SQLz = SQLz & ", ScheduleGroupPaymentTerm"
    SQLz = SQLz & ", SGrpRSLTPmt"               '40
    SQLz = SQLz & ", GroupAddedDate"
    SQLz = SQLz & ", SGrpFAID"
    SQLz = SQLz & ", Fee_Projected"
    SQLz = SQLz & ", POVendor"
    SQLz = SQLz & ", RebateUnitProjected "      '45
    SQLz = SQLz & ", Assignee_Bank"
    SQLz = SQLz & ", Assigned_Indicator "
    SQLz = SQLz & ", AssigneeBankID "
    SQLz = SQLz & ", BPChgLO"
    SQLz = SQLz & ", BPChgLO_Up"                '50
    SQLz = SQLz & ", BPChgLO_Down"
    SQLz = SQLz & ", SGrpSalesBonus"
    SQLz = SQLz & ", FundDate_Projected"        '53
    SQLz = SQLz & ") SELECT DISTINCT AG.ATFID"
    SQLz = SQLz & ", Right(ATFGroupSpecDesc,1) AS SGrpNO"
    SQLz = SQLz & ", Schedule.SchID AS SGrpSchID"
    SQLz = SQLz & ", AG.ATFGroupSpecID"
    SQLz = SQLz & ", AG.UnitsInGroup"           '5
    SQLz = SQLz & ", AG.GroupOEC"
    SQLz = SQLz & ", AG.ExchangeITGroup"
    SQLz = SQLz & ", AG.ExchangeITType"
    SQLz = SQLz & ", AG.ExchangeITMonth "
    SQLz = SQLz & ", AG.ExtendITOptionTerm"     '10
    SQLz = SQLz & ", AG.ExtendedRentAvg"
    SQLz = SQLz & ", AG.MTMRentAvg"
    SQLz = SQLz & ", AG.PrimaryLeaseTerm"
    SQLz = SQLz & ", AG.PrimaryLeaseNotice"
    SQLz = SQLz & ", AG.LeasePymtFreq"          '15
    SQLz = SQLz & ", AG.LeasePaidAdvArr"
    SQLz = SQLz & ", AG.SchRentAvg"
    SQLz = SQLz & ", AG.LRFSchGroup"
    SQLz = SQLz & ", AG.SwapIndex"
    SQLz = SQLz & ", AG.SwapDate"               '20
    SQLz = SQLz & ", AG.SwapRate"
    SQLz = SQLz & ", AG.SwapTermYears"
    SQLz = SQLz & ", AG.SwapYield"
    SQLz = SQLz & ", AG.SwapBonus AS SGrpBonus"
    SQLz = SQLz & ", AG.PLB"                    '25
    SQLz = SQLz & ", 'FMV' AS LeaseType"
    SQLz = SQLz & ", AG.BLCDBasisID"
    SQLz = SQLz & ", 10 AS PerDiemStartsOn"
    SQLz = SQLz & ", AG.AllowanceMethod"
    SQLz = SQLz & ", AG.MPYAllowance"           '30
    SQLz = SQLz & ", AG.CPMile"
    SQLz = SQLz & ", AG.CPMileThreshold"
    SQLz = SQLz & ", AG.CPMileAfter"
    SQLz = SQLz & ", AG.HPYAllowance"
    SQLz = SQLz & ", AG.CPHour"                 '35
    SQLz = SQLz & ", AG.CPHourThreshold"
    SQLz = SQLz & ", AG.CPHourAfter"
    SQLz = SQLz & ", AG.FMSFeeAmount"
    SQLz = SQLz & ", AG.FMSFeeTermMonths"
    SQLz = SQLz & ", AG.Fee2Amount"             '40
    SQLz = SQLz & ", Date() AS GroupAddedDate"
    SQLz = SQLz & ", AG.GrpFAID"
    SQLz = SQLz & ", AG.Fee_Projected"
    SQLz = SQLz & ", U.VendorID"
    SQLz = SQLz & ", U.UTA"                     '45
    SQLz = SQLz & ", 'TBD' AS Assignee"
    SQLz = SQLz & ", 'P' AS Assigned"
    SQLz = SQLz & ", 5649 AS AssigneeBankID"
    SQLz = SQLz & ", 10 AS BPChgLO"
    SQLz = SQLz & ", AG.SwapAdjustment AS BP_Up" '50
    SQLz = SQLz & ", AG.SwapAdjustment AS BP_Down"
    SQLz = SQLz & ", AG.SwapBonus AS SalesBonus"
    SQLz = SQLz & ", AG.FundDate_Projected"      '53
    SQLz = SQLz & " FROM (ATFGroupSpec As AG"
    SQLz = SQLz & " INNER JOIN Schedule ON AG.ATFSchSpecID = Schedule.ATFSchSpecID)"
    SQLz = SQLz & " INNER JOIN (Select ATFGroupSpecID, VendorId, Avg(UTARebate) As UTA From ATFUnitSpec"
    SQLz = SQLz & "  Group By ATFGroupSpecID, VendorId)  AS U"
    SQLz = SQLz & " ON AG.ATFGroupSpecID = U.ATFGroupSpecID"
    SQLz = SQLz & " WHERE AG.GrpFAID= '" & Me.FAID & "';"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True

DeployUnits:
    Dim GroupCntInt As Integer
    Dim UnitCntInt As Integer
    Dim TBDLocation As Long
    On Error Resume Next
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete From [ATF_UnitLevelHash] WHERE ATFID=" & ATFIDx & "")
    DoCmd.SetWarnings True
    On Error GoTo 0
    Dim Rst2 As DAO.Recordset
    SQLz = ""
    SQLz = SQLz & "SELECT SchGrp.SGrpNo, SchGrp.SGrpID, SchGrp.SGrpSchID, AG.ATFGroupSpecID, AG.UnitSpecsInGroup, AG.UnitsInGroup, AG.GrpFAID, ATFUnitSpec.ATFUnitSpecID "
    SQLz = SQLz & "FROM (SchGrp INNER JOIN ATFGroupSpec AS AG ON SchGrp.ATFGroupSpecID = AG.ATFGroupSpecID) INNER JOIN ATFUnitSpec ON AG.ATFGroupSpecID = ATFUnitSpec.ATFGroupSpecID "
    SQLz = SQLz & "WHERE AG.[GrpFAID]='" & Me.FAID & "' "
    SQLz = SQLz & "ORDER BY SchGrp.SGrpNo;"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    Set Rst2 = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    RST.MoveLast
    Rst2.MoveLast
    RST.MoveFirst
    Rst2.MoveFirst
    GroupCntInt = RST.RecordCount
    For i = 1 To RST.RecordCount
        GroupSpecIDx = RST!ATFGroupSpecID
        For ii = 1 To Rst2!UnitSpecsInGroup
            Set qdf = CurrentDb.QueryDefs("ATF_InsertUnit")
            qdf.Parameters("GroupIDx").Value = Rst2!SGrpID
            qdf.Parameters("UnitSpecX").Value = Rst2!ATFUnitSpecID
            UnitCntInt = UnitCntInt + Rst2!UnitsInGroup
            For iii = 1 To Rst2!UnitsInGroup
                MessageUser "Update", "Inserting Group " & RST!SGrpNO & " Units", "Unit: " & iii & " of " & Rst2!UnitsInGroup & "", True
                DoCmd.SetWarnings False
                qdf.Execute
                DoCmd.SetWarnings True
            Next iii
            Rst2.MoveNext
        Next ii

        DoCmd.SetWarnings False
        'Insert Missing TBD Location for ClientGroup
        DoCmd.RunSQL ("INSERT INTO Locations " & _
                      "(ClientGroupID , LocationType, LocationName, Locationsn, LState " & _
                      ") SELECT DISTINCT " & _
                      "ClientGroupID, 'Garage','TBD','TBD','XX' From vw_SixKeys " & _
                      "WHERE ClientGroupId NOT in( Select ClientGroupID FROM Locations WHERE LocationSn='TBD')")

        'Get the TBD Location for This ClientGroup
        TBDLocation = Nz(DLookup("LocationID", "Locations", "LocationSN='TBD' AND ClientGroupID=" & DLookup("ClientGroupID", "vw_SixKeys", "ClientID=" & ClientID)), 0)

        'Set the Temp VIN, UnitNumber OBCNumber and TBD Location for the new units
        DoCmd.RunSQL ("UPDATE Units SET" & _
                    " UnitVIN = 'TBD-' & [Units].[unitId] " & _
                      ",UnitUnitNum = 'TBD-' & [Units].[unitId] " & _
                      ",ObcUnitNumber = 'TBD-' & [Units].[unitId] " & _
                      ",UnitGarageLocationRef = " & TBDLocation & " " & _
                      "WHERE GroupREF=" & RST!SGrpID & ";")

        DoCmd.SetWarnings True

        Call MakeATFRevision(RST!SGrpSchID, "Schedule Inserted into ATLAAS", False, False)
        RST.MoveNext
    Next i
    Rst2.close
    RST.close
Insert_RevNote0:
    Dim Schx, Grpx, LinewBullet As String
    LinewBullet = "&nbsp;&nbsp;<Br />" & Chr(149) & "&nbsp;"
    SQLz = "SELECT Schedule.SchNo, Schedule.SchID, SchMLRef, DateSchAdded " & _
           "From Schedule WHERE SchMLRef=" & Me.MLID & " AND DateSchAdded=#" & Date & "#;"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    RST.MoveLast
    RST.MoveFirst
    For i = 1 To RST.RecordCount
        GroupCntInt = DCount("SGrpSchID", "SchGrp", "SGrpSchId=" & RST!SchID)
        UnitCntInt = DCount("UnitID", "vw_SixKeys", "SchID=" & RST!SchID)
        Schx = "ATLAAS Core imported from Salesforce:" & LinewBullet & "Schedule " & RST!SchNo
        If GroupCntInt > 1 Then Grpx = LinewBullet & "with " & GroupCntInt & " Groups" Else Grpx = LinewBullet & "Group A"
        If GroupCntInt = 1 And UnitCntInt > 1 Then UnitX = LinewBullet & "containing " & UnitCntInt & " units"
        If GroupCntInt = 1 And UnitCntInt = 1 Then UnitX = LinewBullet & "with a single unit"
        If GroupCntInt > 1 And UnitCntInt > 1 Then UnitX = LinewBullet & "containing " & UnitCntInt & " units in all"
        SQLz = "INSERT INTO ATF_ChangeReasons "
        SQLz = SQLz & "( [Action], SchID, SchNo, ChangeDate, Reason, SysUser, FieldName, FieldGroup, NewValue, RevID )"
        SQLz = SQLz & "Values ("
        SQLz = SQLz & "'Schedule Import'"
        SQLz = SQLz & ", " & RST!SchID & ""
        SQLz = SQLz & ",'" & RST!SchNo & "'"
        SQLz = SQLz & ",#" & Now() & "#"
        SQLz = SQLz & ",'" & Schx & Grpx & UnitX & "'"
        SQLz = SQLz & ", '" & FAUserName & "'"
        SQLz = SQLz & ", 'Rev-Notes'"
        SQLz = SQLz & ", 'Schedule Notes'"
        SQLz = SQLz & ", 'Note 0'"
        SQLz = SQLz & ", 0"
        SQLz = SQLz & ");"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
        RST.MoveNext
    Next i
OpenReports:
    'Exit Sub ' remove remark to test
    '*************** This will also send out the emails
    RST.MoveFirst
    Call OpenScheduleControl(DMin("UnitID", "vw_SixKeys", "SchID=" & RST!SchID))
    ' For i = 1 To RST.RecordCount
    '     Call OpenActiveATFReport(RST!SchID, True, True)    'True = send out the emails
    '     RST.MoveNext
    ' Next i
    RST.close
    Set RST = Nothing
    Set qdf = Nothing
    MessageUser "Close"
    DoCmd.close acForm, Me.Form.Name
End Sub

'Private Sub LesseeID_Click()
''    Me.Parent.clientID = Me.LesseeID.Column(0)
''    Me.Parent.ClientGroupID = Me.LesseeID.Column(3)
''    SQLz = "SELECT clientID, clientCompanyName AS CoLessee, ClientShNm AS ShortName, ClientGroupId FROM Clients " & _
''           "WHERE ClientGroupId=" & Me.LesseeID.Column(3) & " ORDER BY clientCompanyName;"
''    Me.CoLesseeID.RowSource = SQLz
''
''    SQLz = "SELECT MLID, MLNo, Client, MLOrig FROM MstrLease " & _
''           "WHERE MLOrig='FA' " & _
''           "AND Client=" & Nz(Me.LesseeID, 0) & " " & _
''           "ORDER BY MLID DESC , MLNo;"
''    Me.MLID.RowSource = SQLz
'End Sub

'Public Sub MakeScheduleSpec_Click()
'    If Nz(Me.NumOfSchedules, 0) = 0 Then
'        MsgBox "Please set the number of Schedules in this Lease.", vbInformation, AtlaasVersion
'        Me.NumOfSchedules.SetFocus
'        Exit Sub
'    End If
'    Dim xi, xCount As Integer
'    Dim TempStr As String
'    xCount = DCount("ATFSchSpecID", "ATFScheduleSpec", "ATFID=" & Me.ATFID)
'    If xCount >= Me.NumOfSchedules Then
'        MsgBox "All of the Schedule Specs within the ATF Spec have been created.", vbInformation, AtlaasVersion
'        Exit Sub
'    Else
'        If xCount > 0 Then LastExitstingRule = CInt(Right(DMax("ATFSchSpecDesc", "ATFScheduleSpec", "ATFID=" & DMax("ATFID", "ATF", "ATFD=" & Me.ATFID)), 1))
'
'        If Me.NumOfSchedules > 1 Then TempStr = " Schedule Specs" Else TempStr = " Schedule Spec"
'        If AskToDoTask("You're about to Insert " & Me.NumOfSchedules & TempStr) = True Then
'            For xi = (1 + LastExitstingRule) To Me.NumOfSchedules
'                SQLz = ""
'                SQLz = SQLz & "INSERT INTO ATFScheduleSpec ("
'                SQLz = SQLz & "[ATFID]"
'                SQLz = SQLz & ",[MLID]"
'                SQLz = SQLz & ",[ATFSchName]"
'                SQLz = SQLz & ",[ATFSchSpecDesc]"
'                SQLz = SQLz & ",[SchFAID]"
'                SQLz = SQLz & ",[SchDescription]"
'                SQLz = SQLz & ",[NumGroupsInSch]"
'                SQLz = SQLz & ",[SchSpecInserted]"
'                SQLz = SQLz & ")"
'                SQLz = SQLz & " Values ("
'                SQLz = SQLz & Me.ATFID
'                SQLz = SQLz & "," & Me.MLID
'                SQLz = SQLz & ",'Schedule " & xi & "'"
'                SQLz = SQLz & ",'Schedule " & xi & "'"
'                SQLz = SQLz & ",'" & Me.FAID & "'"
'                SQLz = SQLz & ",'Enter Sch " & xi & " Desc Here'"
'                SQLz = SQLz & ",1"
'                SQLz = SQLz & ",#" & Date & "#)"
'                DoCmd.SetWarnings False
'                DoCmd.RunSQL (SQLz)
'                DoCmd.SetWarnings True
'            Next xi
'
'        End If
'    End If
'End Sub
Private Sub Refresh_My_Buttons()
    Dim ctrl   As Control
    On Error GoTo Refresh_My_Buttons_Error

    For Each ctrl In Me.Controls
        If ctrl.ControlType = acCommandButton Then ctrl.Requery
    Next ctrl
    On Error GoTo 0
    Exit Sub

Refresh_My_Buttons_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Refresh_My_Buttons of VBA Document Form_ATF_ATFSpec"
End Sub



Private Sub TabCtrl1_Click()
    Exit Sub
End Sub
