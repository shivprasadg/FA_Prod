Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim Ccrl       As Control
Public OldLocationX As Integer
Dim SQLz As String
Dim UserAG As String
Public PrevStatus As String
Dim UserAccess As Integer



Private Sub AltUnitnum_Change()
    If Me.AltUnitnum.Text & "test" = "test" Then
        Me.AltUnitnum.BackColor = RGB(216, 216, 216)
    Else
        Me.AltUnitnum.BackColor = RGB(255, 255, 255)
    End If
End Sub

Private Sub DecodeMyVin_Click()

MessageUser "Open", "Geting Vin", "Just a sec..."
DoCmd.SetWarnings False
DoCmd.RunSQL ("Delete From UnitVindetails Where UnitId=" & Me.UnitID & ";")
DoCmd.SetWarnings True
    Call VinDecoder(Me.UnitID, Me.unitvin, Me.UnitYr, , , Me.ClientID)
    Me.VINDecode.Form.Requery
MessageUser "Close"
End Sub

Private Sub Form_Close()
'Exit Sub
    On Error GoTo FixMe
    If Me.MLOrig = "FA" Then
        If IsDate(Me.UnitActualDelDate) And IsDate(Me.UnitBuildDate) And Me.UnitActualDelDate < Me.UnitBuildDate Then
            MsgBox "Warning! Delivery Date is less than build date", vbCritical
            Me.UnitActualDelDate.SetFocus
        End If
   '     Call NoticeCheck
        Call CheckATFChanges(Me.GroupRef, "Rent Change", Date, "Units Form")
        Call CheckATFChanges(Me.GroupRef, "Residual Change", Date, "Units Form")
        Call CheckATFChanges(Me.GroupRef, "Unit Spec Change", Date, "Units Form")
        Call CheckATFChanges(Me.GroupRef, "Unit Vendor Change", Date, "Units Form")
        Call CheckATFChanges(Me.GroupRef, "ChildPart Change", Date, "Units Form")
    End If
    Exit Sub
FixMe:
    Exit Sub
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Close of VBA Document Form_UnitsForm"
End Sub

Private Sub NoticeCheck()
    If Me.MLOrig = "FA" Then
        If AllUnitsInGroup_Accepted(Me.GroupRef) = True Then Call AllUnitsInGroup_AcceptedMessage(Me.GroupRef)
        If AllUnitsInGroup_PerDiemSet(Me.GroupRef) = True Then Call AllUnitsInGroup_PerDiemSet(Me.GroupRef)
        'Disabled 8/13/2020 per OPS - Pat
        'If AllUnitsInGroup_Invoiced(Me.GroupRef) = True Then Call AllUnitsInGroup_InvoicedMessage(Me.GroupRef)
    End If
End Sub
Private Sub Form_Load()
 On Error GoTo FixMe
    Dim Ccrl   As Control
    UserAG = FAUserAccessGroup
    UserAccess = FAUserAccess
    'Check User's Access - Lock/UnLock controls
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Requried" Then
            If UserAccess >= 6 Or UserAccess Like "*OPS*" Then
                Me(Ccrl.Name).Locked = False
            Else
                Me(Ccrl.Name).Locked = True
            End If
        End If
        If Me(Ccrl.Name).Tag = "ITTOnly" Then
            If UserAG Like "*ITT*" Then Me(Ccrl.Name).Visible = True Else Me(Ccrl.Name).Visible = False
        End If
    Next Ccrl
    If UserAG Like "*TMT*" Or UserAccess > 50 Then Me.LegalDescription.Locked = False Else Me.LegalDescription.Locked = True
    Me.NoUnits.Value = 1
    'Hide or show AddUnits
    If (UserAccess > 11 And UserAG Like "*OPS*") Or UserAccess = 99 Then
        Me.AddUnits.Visible = True
        Me.NoUnits.Visible = True
        Me.Hidden.Visible = True
    Else
        Me.AddUnits.Visible = False
        Me.NoUnits.Visible = False
        Me.Hidden.Visible = False
    End If
    Exit Sub
FixMe:
    MsgBox "There was an Error with the " & Me.Form.Name & " Form in the Form_Load Event-Error Code:" & Err.Number, vbCritical, "ATLAAS Core Error"
    Resume Next
End Sub

Private Sub Form_Current()
On Error GoTo FixMe
    If Nz(Me.AltUnitnum, "") <> "" Then Me.AltUnitnum.BackColor = RGB(255, 255, 255) Else Me.AltUnitnum.BackColor = RGB(216, 216, 216)
'Check if to lock VIN
    Call UnitVINLocked_Click
    If Not IsDate(Me.BLCD) Then
        Me.UnitCOABack.Locked = False
        Me.UnitAcceptDate.Locked = False
    Else
        Me.UnitCOABack.Locked = True
        Me.UnitAcceptDate.Locked = True
    End If
'Set UOM Label and ComboBox
    Call SetUOM(Me.UnitSubType, Me.UnitType, "UOMCbo")
'Get ChildPart cost and OEC totals
    Me.UnitGarageLocationRef.Locked = False
    If Me.Form.NewRecord = False Then
        If Me.UnitStatus < "C" And UserAccess <= 15 Then Me.UnitGarageLocationRef.Locked = True
        Me.OEC = CCur(Nz([UnitFinalCost], 0) + Nz(DSum("childFinalCost", "ChildParts", "[UnitREF]=" & Me.UnitID & ""), 0))
        Me.ChildPartCost = CCur(Nz(Me.OEC, 0) - Nz(Me.UnitFinalCost, 0))
    '    Me.UnitGarageLocationRef.Requery   ' 7/26/2018 LBJ
    '    Me.UnitTitleLocationRef.Requery    ' 7/26/2018 LBJ
    '    Me.UnitRegLocationRef.Requery      ' 7/26/2018 LBJ
    '    Me.DeliveryCbo.Requery             ' 7/26/2018 LBJ
    '    Me.DivisionOwnerID.Requery         ' 7/26/2018 LBJ
        Call MakeLRF
        ' No need to restrict Gargage Location on Not A or B units 4/11/2017 LBJ
        ' If Me.unitstatus > "B" And FAUserAccess <= 14 Then Me.UnitGarageLocationRef.Locked = True
        ' Removed call to SetSubGroup 4/11/2017 LBJ
        ' Call SetSubGroup
    End If
    Call UnitType_Click
    Call UnitSubType_Click
    'Me.Form.Repaint
    Exit Sub
FixMe:
    MsgBox "There was an Error with the " & Me.Form.Name & " Form in the Form_Current Event- Error Code:" & Err.Number, vbCritical, "ATLAAS Core Error"
    Resume Next
End Sub
Public Sub SetUOM(UnitSubTypeX, UnitTypeX, ControlNamex As String)
    Dim UomX, ix, RangeLow, RangeHigh, StepX As Integer
    Dim ListX, MeasureType As String
    Me(ControlNamex).ColumnWidths = "0;0.8"
    If UnitTypeX = "Trailer" Then UomX = 30
    If UnitSubTypeX = "Tank" Then UomX = 37
    If UnitSubTypeX = "Sleeper" Then UomX = 12
    If UnitSubTypeX = "Step Van" Then UomX = 20
    If UnitSubTypeX = "Auto Hauler" Then UomX = 36
    If UnitTypeX = "Straight Truck" Then UomX = 20
    If UnitSubTypeX = "Tank" And UnitTypeX = "Straight Truck" Then UomX = 27
    If UomX = 27 Or UomX = 37 Then StepX = 500 Else StepX = 1
    If UomX <> 0 Then
        RangeLow = DMin("UOMRangeLow", "Lookup_UnitUOM", "UOMGroupID=" & UomX)
        RangeHigh = DMax("UOMRangeHigh", "Lookup_UnitUOM", "UOMGroupID=" & UomX)
        MeasureType = DLookup("UOMPrintLabel", "Lookup_UnitUOM", "UOMGroupID=" & UomX)
        Me(ControlNamex).ControlTipText = "A " & SubUnitTypeX & " " & UnitTypeX & "'s UOM Range is: " & DLookup("UOMLabel", "vw_UnitType_Matrix", "UOMGroupID = " & UomX)

        For ix = RangeLow To RangeHigh Step StepX
            ListX = ListX & ix & ";" & ix & MeasureType & ";"
        Next ix

        Me(ControlNamex).RowSource = ListX
        Me(ControlNamex).Enabled = True
    Else
        Me(ControlNamex).RowSource = ""
        Me(ControlNamex).ControlTipText = "A " & UnitSubTypeX & " " & UnitTypeX & " doesn't have a UOM"
    End If
      Me.UOMCbo.SpecialEffect = 0
End Sub
Private Sub MakeLRF()
    On Error GoTo FixMe
    Dim UnitOEC As Double
    If Nz(Me.UnitFinalCost, 0) = 0 Then
        Exit Sub
    End If
    UnitOEC = CCur(Nz([UnitFinalCost], 0)) + CCur(Nz(DSum("childFinalCost", "ChildParts", "[UnitREF]=" & Me.UnitID & ""), 0))
    If Me.UNITRENT > 0 Then Me.LRFSyn = GetLFR(Me.UNITRENT, UnitOEC)
    If Me.UNITSRENT > 0 Then Me.LRFLo = GetLFR(Me.UNITSRENT, UnitOEC)
    Exit Sub
FixMe:
    MsgBox "There was an Error with the " & Me.Form.Name & " Form - Error Code:" & Err.Number, vbCritical, "ATLAAS Core Error"
    Resume Next
End Sub

'Private Sub SetSubGroup()
' Killed 4/11/2017 LBJ
'    On Error GoTo Fixme
'    SQLz = "SELECT SubType, GrpClass AS Class, ClassID " & _
'           "FROM LookUp_GroupSubTypes Where ClassID =" & Me.UnitType.Column(2) & " Order By SubType;"
'    Me.UnitSubType.RowSource = SQLz
'   Exit Sub
'Fixme:
'    MsgBox "There was an Error with the " & Me.Form.Name & " Form - Error Code:" & Err.Number, vbCritical, "ATLAAS Core Error"
'    Resume Next
'End Sub

Public Sub AddUnits_Click()
    Dim Msgx, StatusX, Committ As String
    Dim NumOfUnits As Integer
    Msgx = "| "
    If IsNull(Me.UnitYr.Value) Then Msgx = Msgx & "UnitYr Missing |"
    If IsNull(Me.UnitMake) Then Msgx = Msgx & "Make Missing |"
    If IsNull(Me.UnitType) Then Msgx = Msgx & "Type Missing |"
    If Me.UnitType = "Tractor" And IsNull(Me.UnitSubType) Then Msgx = Msgx & "SubType Missing |"
    If IsNull(Me.VendorREF) Then Msgx = Msgx & "Vendor Missing |"
    If IsNull(Me.UnitGarageLocationRef) Then Msgx = Msgx & "Garage Missing |"
    If Me.UnitFinalCost.Value <= 0 Then Msgx = Msgx & "Final Cost not set |"
    If Msgx <> "| " Then
        MsgBox "Please Enter All Required Fields" & vbNewLine & Msgx
        Exit Sub
    Else
        If IsNull(Me.NoUnits) Or Me.NoUnits <= 0 Or Me.NoUnits > 550 Then
            MsgBox "Sorry but the 'Units to Add' box must be between 1 and 50", vbExclamation, "You're Outta Control!"
            Me.NoUnits.SetFocus
            Exit Sub
        End If
        If IsNull(Me.UnitsClientID) Then Me.UnitsClientID = Me.ClientID
        If IsNull(Me.ClientGroupID) Then ClientGroupID = DLookup("ClientGroupId", "vw_ClientGroup_MatchedClients", "ClientId=" & Me.ClientID & "")
        If Me.Form.NewRecord = True Then NumOfUnits = Me.NoUnits - 1 Else NumOfUnits = Me.NoUnits
        If Me.UnitStatus & "test" = "test" Then Me.UnitStatus = "B"
        Me.Form.Refresh
        Committ = MsgBox("You're about Insert " & Me.NoUnits & " Units" & vbNewLine & "Into Group: " & SGrpNO & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Adding " & TragetX & "Units?")
        If Committ = vbNo Then Exit Sub
        Dim SQLz, BuildDateX, EstDateX, PDateX As String
        If IsNull(Me.UnitBuildDate) Then BuildDateX = "" Else BuildDateX = Me.UnitBuildDate
        If IsNull(Me.UnitEstDelDate) Then EstDateX = "" Else EstDateX = Me.UnitEstDelDate
        If IsNull(Me.UnitPDStartDate) Then PDateX = "" Else PDateX = Me.UnitPDStartDate
        Me.ClientGroupID = DLookup("ClientGroupId", "ClientGroup_MatchedClients", "[ClientId]=" & Me.ClientID & "")
        StatusX = "B"
        If Me.MLOrig <> "FA" Then StatusX = "T"
        If Me.MLOrig <> "FA" And Me.MLLessor = "FA LLC" Then StatusX = "R"

SQLString:
        SQLz = "INSERT INTO Units (UnitBuildDate, UnitEstDelDate, CofASigner, POASigner, UnitFinalCost, UnitDeliveryLocationRef, UnitGarageLocationRef, UnitTitleLocationRef, " & _
               "UnitRegLocationRef, UnitIRP, UnitTaxExempt, UnitTaxNotes, UnitPDStartDate, UnitRent, UnitsRent, UnitPDRent, MTMRent, FirstExtRent, UnitTAX, UnitSalesResid, " & _
               "UnitYr, UnitMake, UnitModel, UnitType, UnitSubType, MaintProv, UnitDel, UnitDep, UnitDepdt, Axle, "

        If Nz(Me.UnitUOM, 0) <> 0 Then SQLz = SQLz & "UnitUOM, "
        If Nz(Me.UnitConfiguration, "") <> "" Then SQLz = SQLz & "UnitConfiguration, "
        If Nz(Me.UnitAxleProperty, "") <> "" Then SQLz = SQLz & "UnitAxleProperty, "
        If Nz(Me.UnitDescriptor, "") <> "" Then SQLz = SQLz & "UnitDescriptor, "

        SQLz = SQLz & "UnitPilot"
        SQLz = SQLz & ", UnitVINLocked"
        SQLz = SQLz & ", UnitPLB"
        SQLz = SQLz & ", UnitDocAccpt"
        SQLz = SQLz & ", UnitTitlingTrust"
       ' Sqlz = Sqlz & ", PRINTCOFA"
        SQLz = SQLz & ", PRINTPOFA"
      '  Sqlz = Sqlz & ",PNUNIT"
        SQLz = SQLz & ",GroupRef"
        SQLz = SQLz & ", VendorRef"
        SQLz = SQLz & ", ClientGroupID"
        SQLz = SQLz & ", UnitStatus"
        SQLz = SQLz & ", ClientID ) "
        SQLz = SQLz & "SELECT "
        SQLz = SQLz & "'" & BuildDateX & "'"
        SQLz = SQLz & ", '" & EstDateX & "'"
        SQLz = SQLz & ", " & Nz(Me.CofASigner, 0) & ""
        SQLz = SQLz & ", " & Nz(Me.POASigner, 0)
        SQLz = SQLz & ", " & Me.UnitFinalCost & ""
        SQLz = SQLz & ", " & Nz(Me.UnitDeliveryLocationRef, 0) & ""
        SQLz = SQLz & ", " & Nz(Me.UnitGarageLocationRef, 0) & ""
        SQLz = SQLz & ", " & Nz(Me.UnitTitleLocationRef, 0) & ""
        SQLz = SQLz & ", " & Nz(Me.UnitRegLocationRef, 0) & ""
        SQLz = SQLz & ", " & Me.UnitIRP & ""
'        SQLz = SQLz & ", '" & Me.UnitIRPAcct & "'"
        SQLz = SQLz & ", " & Me.UnitTaxExempt & ""
        SQLz = SQLz & ", '" & Me.UnitTaxNotes & "'"
        SQLz = SQLz & ", '" & PDateX & "'"
        SQLz = SQLz & ", " & Nz(Me.UNITRENT, 0) & ""
        SQLz = SQLz & ", " & Nz(Me.UNITSRENT, 0) & ""
        SQLz = SQLz & ", " & Nz(Me.UNITPDRENT, 0) & ""
        SQLz = SQLz & ", " & Nz(Me.MTMRent, 0) & ""
        SQLz = SQLz & ", " & Nz(Me.FirstExtRent, 0) & ""
        SQLz = SQLz & ", " & Nz(Me.UnitTax, 0) & ""
        SQLz = SQLz & ", " & Nz(Me.UnitSalesResid, 0) & ""
        SQLz = SQLz & ", " & Me.UnitYr & ""
        SQLz = SQLz & ", '" & Me.UnitMake & "'"
        SQLz = SQLz & ", '" & Me.UnitModel & "'"
        SQLz = SQLz & ", '" & Me.UnitType & "'"
        SQLz = SQLz & ", '" & Me.UnitSubType & "'"
        SQLz = SQLz & ", '" & Me.MAINTPROV & "'"
        SQLz = SQLz & ", '" & Me.UnitDel & "'"
        SQLz = SQLz & ", " & Me.UnitDep & ""
        SQLz = SQLz & ", #" & Nz(Me.UnitDepdt, #1/1/2050#) & "#"
        SQLz = SQLz & ", '" & Me.Axle & "'"
        If Nz(Me.UnitUOM, 0) <> 0 Then SQLz = SQLz & "," & Me.UnitUOM & " "
        If Nz(Me.UnitConfiguration, "") <> "" Then SQLz = SQLz & ",'" & Me.UnitConfiguration & "'"
        If Nz(Me.UnitAxleProperty, "") <> "" Then SQLz = SQLz & ",'" & Me.UnitAxleProperty & "'"
        If Nz(Me.UnitDescriptor, "") <> "" Then SQLz = SQLz & ",'" & Replace(Me.UnitDescriptor, "'", "''", 1) & "'"
        SQLz = SQLz & ",0, 0, 0, -1, 0, 0, " & Me.GroupRef & ", " & Me.VendorREF & ", " & Me.ClientGroupID & ", '" & StatusX & "', " & Me.ClientID & ";"

        DoCmd.SetWarnings False
        For i = 1 To NumOfUnits
            DoCmd.RunSQL (SQLz)
        Next i
    End If
    DoCmd.SetWarnings True
End Sub

Private Sub CofaSign_GotFocus()
    Me.CofaSign.Requery
End Sub

Private Sub DeliveryCbo_GotFocus()
    Me.DeliveryCbo.Requery
End Sub

Private Sub ObcUnitNumber_DblClick(Cancel As Integer)
    Me.ObcUnitNumber.Locked = False
End Sub

Private Sub OffLeaseInd_Change()
    If OffLeaseInd.Text = "Actual" Or OffLeaseInd.Text = "Projected" Then
        If OffLeaseInd.Text = "Actual" And Me.unitoffleasedt > Date Then
            Dim Committ As Integer
            Committ = MsgBox("You're setting the Off Lease Indicator to 'Actual' " & vbNewLine & "When the Off Lease date has not yet passed." & vbNewLine & "Do you want to still make the change?", vbQuestion + vbYesNo + vbCritical, "Mark a future Off lease date as actual?")
            If Committ <> vbYes Then Me.OffLeaseInd = ValueX
        End If
    Else
        Me.OffLeaseInd = ValueX
    End If
End Sub

Private Sub OffLeaseInd_GotFocus()
    If Not IsNull(Me.ActiveControl) Then ValueX = Me.ActiveControl
End Sub

Private Sub OpenUnitsGroup_Click()
    DoCmd.OpenForm "GroupForm", , , "SGrpID = " & Me.SGrpID
End Sub

Private Sub RequiredLabel_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Requried" Then
            Me(Ccrl.Name).SpecialEffect = 0
            Me(Ccrl.Name).FontBold = True
        End If
    Next Ccrl
End Sub

Private Sub ClonedLabel_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Cloned" Then
            Me(Ccrl.Name).SpecialEffect = 0
        End If
    Next Ccrl
End Sub

Private Sub requiredOutbox_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Requried" Or Me(Ccrl.Name).Tag = "Cloned" Then
            Me(Ccrl.Name).SpecialEffect = 0
        End If
    Next Ccrl
End Sub

Private Sub AddTitleLogEntry_Click()
    Dim Response, StateX As String
    If Me.UnitStatus = "T" Or Me.UnitStatus = "S" Or Me.UnitStatus = "R" Or Me.UnitStatus = "PS" Then
        Response = MsgBox("You are attempting to add a Log for a Unit that may not belong to " & vbNewLine & "Fleet Advantage. This Unit may be a Purchase Lease Back or in Remarketing" & vbNewLine & "Do you want to proceed?", vbQuestion + vbYesNo, "Not a FA unit")
    Else
        Response = MsgBox("You're about to ADD a Title Log Record, Are you Sure?", vbQuestion + vbYesNo, "New Record")
    End If
    If Response = vbYes Then
        DoCmd.close acForm, "OPSTitleLogEntry"
        DoCmd.OpenForm "OPSTitleLogEntry", , , , acFormAdd
        Forms!OPSTitleLogEntry.UnitRef = Me.UnitID
        Forms!OPSTitleLogEntry.USERID = FaUserName
        Forms!OPSTitleLogEntry.TitleNumber = Me.UnitTitleNumber
        If IsNull(Me.UnitRegLocationRef) Or Me.UnitRegLocationRef = 0 Then
            On Error Resume Next
            StateX = Nz(Forms!Unitsform!UnitTitleLogForm.Form.TitleState, "")
CheckStatus:
            On Error GoTo 0
        Else
            StateX = Nz(DLookup("LState", "Locations", "LocationID =" & Me.UnitRegLocationRef), "")
        End If
        If Len(StateX) > 1 Then Forms!OPSTitleLogEntry.TitleState = StateX
    Else
        Exit Sub
    End If
End Sub

Private Sub UnitCOABack_AfterUpdate()
    If Me.MLOrig = "FA" Then
        If IsDate(Me.UnitCOABack) And (Me.UnitStatus) = "B" Then Me.UnitStatus.Value = "A"
        If Not IsDate(Me.UnitCOABack) And Not IsDate(Me.UnitCofACopy) And (Me.UnitStatus) = "A" Then Me.UnitStatus.Value = "B"
    End If
End Sub

Private Sub UnitCofACopy_AfterUpdate()
    If Me.MLOrig = "FA" Then
        If IsDate(Me.UnitCofACopy) And (Me.UnitStatus) = "B" Then Me.UnitStatus.Value = "A"
        If Not IsDate(Me.UnitCOABack) And Not IsDate(Me.UnitCofACopy) And (Me.UnitStatus) = "A" Then Me.UnitStatus.Value = "B"
    End If
End Sub
Private Sub UnitAcceptDate_AfterUpdate()
    If Me.MLOrig = "FA" Then
        If IsDate(Me.UnitCofACopy) And Me.UnitStatus = "B" Then Me.UnitStatus.Value = "A"
        If IsDate(Me.UnitCOABack) And Me.UnitStatus = "B" Then Me.UnitStatus.Value = "A"
    End If
End Sub

Private Sub UnitInServ_DblClick(Cancel As Integer)
    Me.UnitInServ.Locked = False
End Sub

Private Sub UnitMake_Click()
    Me.UnitModel.Requery
End Sub

Private Sub unitoffleasedt_AfterUpdate()
 If IsDate(Me.Units_RMKT_Sub!ParkDT) Then UnitStatus = "R"
End Sub

Private Sub UnitStatus_Click()
    If Me.UnitStatus = "A" Then
        If IsNull(Me.UnitCOABack) And IsNull(UnitCofACopy) Then
            MsgBox "Thus unit cannot be set to 'A' without a COA-Back or a COA-Copy date", vbCritical, "Not yet an Accepted Unit"
            Me.Form.Undo
        End If
    End If
    'If Me.unitstatus = "TR" Then
    '    Me.Units_RMKT_Sub.Form.Visible = True
    '    Me.Units_RMKT_Sub!FABrkComm = 0
    '    Me.Units_RMKT_Sub.Form.Refresh
    'End If
End Sub

Private Sub UnitStatus_GotFocus()
PrevStatus = Me.UnitStatus
'If DLookup("[Access Level]", "TBL_Employees", "[User ID] = FAUserName()") = 0 Then
    If FAUserAccessLevel >= 6 Or FAUserAccessGroup Like "*OPS*" Then
        Me.UnitStatus.Locked = False
    Else
        If Me.UnitStatus = "R" Or Me.UnitStatus = "S" Then
            Me.UnitStatus.Locked = True
        End If
    End If
End Sub

Private Sub UnitSubType_Click()
'  SQLz = "SELECT UnitConfig, UnitConfigDesc, UnitConfigID, UnitSubTypeID " & _
   "FROM LookUp_UnitConfiguration Where UnitSubTypeID =" & Me.UnitSubType.Column(2) & " Order By UnitConfig;"
    Me.UnitConfiguration.Requery
End Sub

Private Sub UnitTASent_AfterUpdate()
    If IsDate(Me.UnitTASent) Then
        If DLookup("EntryReasonId", "UnitTitleLog", "UnitRef=" & Me.UnitID) = 18 Then
            MsgBox "A Title Application Date has already be set for Unit " & Me.UnitID & "." & vbNewLine & vbNewLine & "If you need to send for a new or corrected title, " & vbNewLine & "Please enter a new title log below." & vbNewLine & vbNewLine & "Date will revert (not saved)", vbCritical, "Title App Already set"
            Me.Undo
            Exit Sub
        End If
        SQLz = ""
        SQLz = SQLz & "INSERT INTO UnitTitleLog ( UnitRef,  Reason,  UnitTitleLogDate,  EntryTypeId,  EntryReasonId,  ExpectedReturn,  UserID )"
        SQLz = SQLz & "SELECT " & Me.UnitID & ", 'Title Application Sent', #" & Me.UnitTASent & "#, 2, 18, True, '" & FaUserName & "';"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
        Me.UnitTitleLogForm.Requery
    End If
End Sub

Private Sub UnitTitleLocationRef_GotFocus()
    Me.UnitTitleLocationRef.Requery
End Sub

Private Sub UnitRegLocationRef_GotFocus()
    Me.UnitRegLocationRef.Requery
End Sub

Private Sub ClearFilters_Click()
    On Error GoTo FixMe
    Me.Form.Filter = "SchID= " & Me.SchID & ""
    Me.Form.FilterOn = True
FixMe:
    If IsNull(Me.SchID) Then MsgBox "There was an error in clearing the filter" & vbNewLine & "Please go to the Previous Record and then Press 'Clear Filters'", , "No SchId"
End Sub

Private Sub CloseMe_Click()

    If Me.UnitStatus <> "A" Or Me.UnitStatus <> "B" Then
        DoCmd.close acForm, "UnitsForm", acSaveYes
        Exit Sub
    End If
    Dim Msgx   As String
    Msgx = "| "
    If IsNull(Me.UnitYr.Value) Then Msgx = Msgx & "UnitYr Missing |"
    If IsNull(Me.UnitMake) Then Msgx = Msgx & "UnitMake Missing |"
    If IsNull(Me.UnitType) Then Msgx = Msgx & "UnitType Missing |"
    If IsNull(Me.VendorREF) Then Msgx = Msgx & "Vendor Missing |"
    If UnitStatus = "A" Or UnitStatus = "B" Then
        If Me.UnitFinalCost.Value <= 0 Then Msgx = Msgx & "Final Cost not set |"
    End If
    If Msgx <> "| " Then
        Me.Undo
        MsgBox "Required fields are missing" & vbNewLine & Msgx & vbNewLine & "No Changes Made!", vbCritical, "Required Fields Missing"
        DoCmd.close acForm, Me.Form.Name, acSaveNo
        Exit Sub
    End If
    If (IsNull(Me.UnitGarageLocationRef) Or (Me.UnitGarageLocationRef) = 0) And (Me.UnitStatus) = "A" Then
        MsgBox "Please update Garage Location before you can close this record", vbCritical, "Location Fields Missing"
        Exit Sub
    End If
    If Me.Form.Dirty = True Then
        Select Case MsgBox("Do you want to save your changes to the Unit?", vbYesNo)
        Case vbYes
            DoCmd.close acForm, "UnitsForm", acSaveYes
        Case vbNo
            Me.Undo
            DoCmd.close acForm, "UnitsForm", acSaveNo
        End Select
    Else
        DoCmd.close acForm, "UnitsForm", acSaveYes
    End If
End Sub



Private Sub UnitDocAccpt_DblClick(Cancel As Integer)
    DoCmd.OpenForm "OpsPassword"
End Sub

Private Sub UCCC_AfterUpdate()
    If Me.AddUnits.Visible = True Or IsNull(Me.UCCC) Or checkCostacctcd() = 1 Then Exit Sub Else DoCmd.OpenForm "UCCCCHDate"
    Me.Requery
End Sub

Private Sub UCCC_Dirty(Cancel As Integer)
    Dim FADB   As DAO.DataBase
    Dim UCCClogDet As DAO.Recordset
    Set FADB = CurrentDb
    Set UCCClogDet = FADB.OpenRecordset("CostAcctChangeTTemp", dbOpenDynaset, dbSeeChanges)
    If Me.AddUnits.Visible = True Or IsNull(Me.UCCC) Or checkCostacctcd() = 1 Then
        Exit Sub
    Else
        UCCClogDet.AddNew
        UCCClogDet("Unitref").Value = Forms!Unitsform!Text373
        UCCClogDet("OldUCCC").Value = Forms!Unitsform!UCCC
        UCCClogDet("User").Value = FaUserName
        UCCClogDet.Update

    End If
End Sub

Private Sub UnitGarageLocationRef_GotFocus()
    Me.UnitGarageLocationRef.Requery
    OldLocationX = Nz(Me.UnitGarageLocationRef, 0)
End Sub

Private Sub UnitGarageLocationRef_AfterUpdate()
    If OldLocationX <> 0 And OldLocationX <> Me.UnitGarageLocationRef.Value Then
        Dim Committ As Integer
        Committ = MsgBox("You're about to change the Garage Location of this Unit" & vbNewLine & "Are you sure?", vbQuestion + vbYesNo + vbCritical, "Location change")
        If Committ = vbYes Then
            Dim ReasonX As String
            Dim AsOfDatex As Date
InputReason:
            ReasonX = InputBox("What is the Reason the unit is moving?" & vbNewLine & "This will be used for the Unit's Location History", "Reason for move...", "Client Requested")
            If ReasonX = "" Then GoTo InputReason
InputDate:
            On Error GoTo InputDate
            AsOfDatex = InputBox("What is the Actual Date the Unit Moved?" & vbNewLine & "please keep correct date format mm/dd/yyyy", "Date this Unit Actually moved.", Date)
            If Not IsDate(AsOfDatex) Then GoTo InputDate
            On Error GoTo FixMe
            SQLz = "INSERT INTO LocationChangeHistory ( AssetId, OldLocationId, NewLocationId, LocationChangeDate, LocationType, Reason, GroupId, FAUserName ) " & _
                   "SELECT " & Me.UnitID & " AS AssetId, " & OldLocationX & " AS OldLocationId, " & Me.UnitGarageLocationRef.Value & " AS NewLocationId, #" & AsOfDatex & "# AS LocationChangeDate, 'Garage'   AS LocationType, '" & ReasonX & "' as Reason, " & Me.SGrpID & "  AS GroupId, FAUserName();"
            DoCmd.SetWarnings False
            DoCmd.RunSQL (SQLz)
            DoCmd.SetWarnings True
        Else
            Me.UnitGarageLocationRef = OldLocationX
            Me.UnitGarageLocationRef.Requery
            OldLocationX = Nz(Me.UnitGarageLocationRef, 0)
        End If
        OldLocationX = Me.UnitGarageLocationRef
    End If
FixMe:
End Sub

Private Sub UnitGarageLocationRef_Dirty(Cancel As Integer)
    Exit Sub
End Sub

Private Sub UnitStatus_BeforeUpdate(Cancel As Integer)
    If Not IsNull(Me.UnitCOABack) And Not IsNull(Me.UnitAcceptDate) And Not ((Me.MLOrig) = "Portfolio" Or (Me.MLOrig) = "Other") And ((Me.UnitStatus) = "T" Or (Me.UnitStatus) = "B") Then
        Cancel = True
        Me.UnitStatus.Undo
    Else
        If Not IsNull(Me.UnitCOABack) And Not IsNull(Me.UnitAcceptDate) And ((Me.MLOrig) = "Portfolio" Or (Me.MLOrig) = "Other") And (Me.UnitStatus) = "A" Then
            Cancel = True
            Me.UnitStatus.Undo
        End If
    End If

    If (IsNull(Me.UnitGarageLocationRef) Or (Me.UnitGarageLocationRef) = 0) And (Me.UnitStatus) = "A" Then
        MsgBox "Please add Garage Location and then update the Unit Status"
        Cancel = True
        Me.UnitStatus.Undo
    End If
    If IsNull(Me.Units_RMKT_Sub!ParkDT) And Me.UnitStatus = "P" Then
        MsgBox "You must have a Parked Date set before you can Park ('P') a unit"
        Cancel = True
        Me.UnitStatus.Undo
    End If
    If PrevStatus = "C" Then
        If Me.UnitStatus <> "C" Then
            MsgBox "Once a Unit is a Casualty ('C'), it will remain a 'C' Status Unit, even when Parked or Off Leased."
        End If
        Me.UnitStatus = "C"
    End If

End Sub

Private Sub UnitStatus_DblClick(Cancel As Integer)
    If Me.AllowAdditions = True And Me.AllowEdits = True And Me.AllowDeletions = True Then
        Exit Sub
    Else
        DoCmd.OpenForm "unlock unit form"
    End If
End Sub

Private Sub UnitTitleNumber_GotFocus()
    If DLookup("[Access Level]", "TBL_Employees", "[User ID] = FAUserName()") = 0 Then
        Me.UnitTitleNumber.Locked = False
    End If
End Sub

Private Sub UnitType_Click()
'Call SetSubGroup
' SQLz = "SELECT UnitSubType, UnitSubTypeDesc, UnitSubTypeID, UnitTypeID " & _
  "FROM LookUp_UnitSubTypes Where UnitTypeID =" & Me.UnitType.Column(2) & " Order By UnitSubType;"
    Me.UnitSubType.Requery
    '  Me.UnitSubType.RowSource = SQLz
End Sub

Private Sub UnitVINLocked_Click()
    If Me.UnitVINLocked = -1 Then
        Dim Ccrl As Control
        For Each Ccrl In Me.Controls
            If Me(Ccrl.Name).Tag = "Requried" Then Me(Ccrl.Name).Locked = True
        Next Ccrl
    Else
        For Each Ccrl In Me.Controls
            If Me(Ccrl.Name).Tag = "Requried" Then Me(Ccrl.Name).Locked = False
        Next Ccrl
    End If
End Sub

Private Sub UOMCbo_GotFocus()
Call SetUOM(Me.UnitSubType, Me.UnitType, "UOMCbo")
End Sub

Private Sub VendorREF_DblClick(Cancel As Integer)
'wire instruction Report VendorPmtsgrpInstRpt
End Sub
