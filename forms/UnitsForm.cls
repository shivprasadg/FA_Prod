Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim Ccrl       As Control
Public OldLocationLng As Long
Dim SQLz       As String
Dim UserAG     As String
Public PrevStatus As String
Public PrevOffLease As Date
Private ranCurrent As Boolean
Dim UserAccess As Integer



Private Sub AltUnitnum_Change()
    If Me.AltUnitnum.Text & "test" = "test" Then
        Me.AltUnitnum.BackColor = RGB(216, 216, 216)
    Else
        Me.AltUnitnum.BackColor = RGB(255, 255, 255)
    End If
End Sub

Private Sub LocationParkedID_AfterUpdate()
    If Me.LocationParkedID.Value = Null Or Me.LocationParkedID = Null Then Exit Sub
    If OldLocationLng <> Me.LocationParkedID.Value And DLookup("ClientGroupId", "Locations", "LocationID=" & Me.LocationParkedID.Value & "") = 771 Then
        Dim Committ As Integer
        Committ = MsgBox("Do you want to create a Move Order to This FA RMKT Location" & vbNewLine & "YES / NO?", vbQuestion + vbYesNo + vbCritical, "Location change")
        If Committ = vbYes Then
            If OldLocationLng = 0 Then OldLocationLng = Nz(Me.UnitGarageLocationRef, 0)
            Call ParkedLocationMove(OldLocationLng, Me.LocationParkedID.Value, Me.UnitID)
            Me.UnitsTabControl.Pages.Item(7).SetFocus
            Me.MoveOrdersForm.Form.Visible = True
            Me.MoveOrdersForm.Form.Detail.Visible = True
            Me.MoveOrdersForm.Requery
            If Me.unitstatus Like "E*" Then Me.unitstatus = "EM"
        Else
            If Nz(OldLocationLng, 0) > 0 And DLookup("ClientGroupId", "Locations", "LocationID=" & Me.OldLocationLng & "") <> 771 Then Me.LocationParkedID = OldLocationLng
            If Nz(OldLocationLng, 0) = 0 Then Call MsgBox("This unit is now assigned to 'RMKT' Parked Location," & vbNewLine & "WITHOUT a move order.", vbCritical, "No Move Order set")
            Me.LocationParkedID.Requery
        End If
        OldLocationLng = Nz(Me.LocationParkedID, 0)
    End If
    ' Call SyncUnitsToLocationDivisions
FixMe:
End Sub

Private Sub LocationParkedID_DblClick(Cancel As Integer)
    DoCmd.close acForm, "LocationAddressesEdit"
    DoCmd.OpenForm "LocationAddressesEdit", acNormal, , , , , DLookup("ClientGroupID", "Locations", "LocationID=" & LocationParkedID & "") & ":" & LocationParkedID
End Sub

Private Sub LocationParkedID_GotFocus()
    OldLocationLng = Nz(Me.LocationParkedID, 0)
End Sub

Private Sub DecodeMyVin_Click()

    MessageUser "Open", "Geting Vin", "Just a sec..."
    '  DoCmd.SetWarnings False
    '  DoCmd.RunSQL ("Delete From UnitVindetails Where UnitId=" & Me.UnitID & ";")
    ' DoCmd.SetWarnings True
    Call VinDecoder(Me.UnitID, Me.UnitVIN, Me.UnitYr, , , Me.clientID)
    'Call SetEngineType(True)
    SQLz = "SELECT * from UnitVinDetails Where UnitId =" & Me.UnitID & ";"
    Me.VINDecode.Form.RecordSource = SQLz
    MessageUser "Close"
End Sub

Private Sub DeliveryCbo_AfterUpdate()
    Call SyncUnitsToLocationDivisions
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
    Call MakePaymentAdjustment(Me.GroupRef)
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

Private Sub Form_DataChange(ByVal Reason As Long)
    MsgBox "data change"
End Sub

Private Sub Form_Load()
    'On Error GoTo Fixme
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
        If Me(Ccrl.Name).Tag = "ITTOnly" Or Me(Ccrl.Name).Tag = "NoLook" Then
            If UserAG Like "*ITT*" Then Me(Ccrl.Name).Visible = True Else Me(Ccrl.Name).Visible = False
            If UserAG Like "*BDEX*" Then Me(Ccrl.Name).Visible = False Else Me(Ccrl.Name).Visible = True
        End If
    Next Ccrl
    If UserAG Like "*TMT*" Or UserAccess > 50 Then Me.LegalDescription.Locked = False Else Me.LegalDescription.Locked = True
    Me.NoUnits.Value = 1
    'Hide or show AddUnits
    If UserAG Like "*OPS*" Or UserAccess = 99 Then
        Me.AddUnits.Visible = True
        Me.NoUnits.Visible = True
    Else
        Me.AddUnits.Visible = False
        Me.NoUnits.Visible = False
    End If
    On Error GoTo 0
    Exit Sub
FixMe:
    MsgBox "There was an Error with the " & Me.Form.Name & " Form in the Form_Load Event-Error Code:" & Err.Number, vbCritical, "ATLAAS Core Error"
End Sub

Private Sub Form_Current()
    On Error GoTo FixMe
    If ranCurrent = True Then Exit Sub     ' prevents double execution
    ranCurrent = True
    Call SetSubData(Me.UnitsTabControl.Value)  ' Locations Tab on load
    If Nz(Me.AltUnitnum, "") <> "" Then Me.AltUnitnum.BackColor = RGB(255, 255, 255) Else Me.AltUnitnum.BackColor = RGB(216, 216, 216)
    If Nz(Me.UnitIFAmt, 0) > 0 And FAUserAccess < 99 Then Me.VendorCost.Locked = True Else Me.VendorCost.Locked = False
    If Nz(Me.UnitIFAmt, 0) > 0 And FAUserAccess < 99 Then Me.UnitFinalCost.Locked = True Else Me.UnitFinalCost.Locked = False
    If Not IsDate(Me.BLCD) Then
        Me.UnitCOABack.Locked = False
        Me.UnitAcceptDate.Locked = False
    Else
        Me.UnitCOABack.Locked = True
        Me.UnitAcceptDate.Locked = True
    End If
    ' check if there are Title notes:
    Dim TitleLogx As Integer
    Me.TitleNotesNotifaction = ""

    TitleLogx = DCount("UnitRef", "UnitNotes", "uNote Like '*Title*' AND UnitRef=" & Me.UnitID)
    If TitleLogx = 1 Then Me.TitleNotesNotifaction = "There is <b>1</b> note about the <b>Title</b>, for this Unit<br /><i>(see Unit Note's Tab)"
    If TitleLogx > 1 Then Me.TitleNotesNotifaction = "There are<b> " & TitleLogx & "</b> notes about the <b>Title</b>, for this Unit<br /><i>(see Unit's Notes Tab)"

    Me.TitleAppSent = Nz(DMax("UnitTitleLogDate", "UnitTitleLog", "EntryReasonId= 18 and UnitRef=" & Me.UnitID & ""), "")
    Me.TitleRecvd = Nz(DMax("UnitTitleLogDate", "UnitTitleLog", "EntryTypeId=1 and UnitRef=" & Me.UnitID & ""), "")
    Me.CurrTitleNumber = Nz(DLookup("TitleNumber", "UnitTitleLog", "UnitTitleLogID=" & Nz(DMax("UnitTitleLogID", "UnitTitleLog", "UnitRef=" & Me.UnitID & " and Nz(TitleNumber,'') <> ''"), 0)), "No Title #")
    ' Check if ther is another unit from a PREVIOUS Lease
    Dim pAssetID As Long
    pAssetID = Nz(DMin("AssetID", "vw_ReLeasedUnits_LastTitleHistory", "VIN='" & Me.UnitVIN & "' AND AssetID <>" & Me.UnitID), 0)
    Me.ReleasedAssetID.Visible = False
    Me.ReleasedAssetIDLabel.Visible = False
    If pAssetID > 0 Then
        If pAssetID < Me.UnitID Then Me.ReleasedAssetIDLabel.caption = "AssetID with Custodian:" _
            Else Me.ReleasedAssetIDLabel.caption = "This Unit Was Re-Leased As AssetID:"
        Me.ReleasedAssetID = pAssetID
        Me.ReleasedAssetID.Visible = True
        Me.ReleasedAssetIDLabel.Visible = True
    End If
    'Get ChildPart cost and OEC totals

    If Me.Form.NewRecord = False Then
        If Me.unitstatus < "C" And UserAccess <= 15 Then Me.UnitGarageLocationRef.Locked = True
        Me.OEC = CCur(Nz([UnitFinalCost], 0) + Nz(DSum("childFinalCost", "ChildParts", "[UnitREF]=" & Me.UnitID & ""), 0))
        Me.ChildPartCost = CCur(Nz(Me.OEC, 0) - Nz(Me.UnitFinalCost, 0))
        Call MakeLRF
        ' No need to restrict Gargage Location on Not A or B units 4/11/2017 LBJ
        ' If Me.unitstatus > "B" And FAUserAccess <= 14 Then Me.UnitGarageLocationRef.Locked = True
        ' Removed call to SetSubGroup 4/11/2017 LBJ
        ' Call SetSubGroup
    End If
    'Check if to lock VIN
    Call UnitVINLocked_Click
    Call UnitType_Click
    Call UnitSubType_Click
    'Set UOM Label and ComboBox
    Call SetUOM(Me.UnitSubType, Me.UnitType, "UOMCbo")
    Me.UnitSouceInfo = "<B>" & UCase([ClientShNm]) & "</b> Lease # <b>" & [MLNo] & " </b>Sch.<b> " & [SchNo] & " </b>Grp.<b>  " & [SGrpNO] & "</b>"
    On Error GoTo 0
    Exit Sub
FixMe:
    MsgBox "There was an Error with the " & Me.Form.Name & " Form in the Form_Current Event- Error Code:" & Err.Number, vbCritical, "ATLAAS Core Error"
    On Error GoTo 0
End Sub
Public Sub SetUOM(UnitSubTypeX, UnitTypeX, ControlNamex As String)
    Dim UomX, ix, RangeLow, RangeHigh, StepX As Integer
    Dim ListX, MeasureType As String
    Me(ControlNamex).ColumnWidths = "0;0.8"
    If UnitTypeX = "Trailer" Then UomX = 30
    If UnitSubTypeX = "Tank" Then UomX = 37
    If UnitSubTypeX = "Daycab" Then UomX = 12
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

Public Sub AddUnits_Click()
    Dim Msgx, StatusX, Committ As String
    Dim NumOfUnits As Integer
    Msgx = "| "
    If IsNull(Me.UnitYr.Value) Then Msgx = Msgx & "UnitYr Missing |"
    If IsNull(Me.UnitMake) Then Msgx = Msgx & "Make Missing |"
    If IsNull(Me.UnitType) Then Msgx = Msgx & "Type Missing |"
    If Me.UnitType = "Tractor" And IsNull(Me.UnitSubType) Then Msgx = Msgx & "SubType Missing |"
    If IsNull(Me.VendorRef) Then Msgx = Msgx & "Vendor Missing |"
    If IsNull(Me.UnitGarageLocationRef) Then Msgx = Msgx & "Garage Missing |"
    If Me.MLOrig = "FA" And Me.UnitFinalCost.Value <= 0 Then Msgx = Msgx & "Final Cost not set |"
    If Msgx <> "| " Then
        MsgBox "Please Enter All Required Fields" & vbNewLine & Msgx
        Exit Sub
    Else
        If IsNull(Me.NoUnits) Or Me.NoUnits <= 0 Or Me.NoUnits > 550 Then
            MsgBox "Sorry but the 'Units to Add' box must be between 1 and 500", vbExclamation, "You're Outta Control!"
            Me.NoUnits.SetFocus
            Exit Sub
        End If
        If IsNull(Me.UnitsClientID) Then Me.UnitsClientID = Me.clientID
        If IsNull(Me.ClientGroupID) Then ClientGroupID = DLookup("ClientGroupId", "vw_ClientGroup_MatchedClients", "ClientId=" & Me.clientID & "")
        If Me.Form.NewRecord = True Then NumOfUnits = Me.NoUnits - 1 Else NumOfUnits = Me.NoUnits
        If Me.unitstatus & "test" = "test" Then Me.unitstatus = "B"
        Me.Form.Refresh
        Committ = MsgBox("You're about Insert " & Me.NoUnits & " Units" & vbNewLine & "Into Group: " & SGrpNO & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Adding " & TragetX & "Units?")
        If Committ = vbNo Then Exit Sub
        Dim SQLz, BuildDateX, EstDateX, PDateX As String
        If IsNull(Me.UnitBuildDate) Then BuildDateX = "" Else BuildDateX = Me.UnitBuildDate
        If IsNull(Me.UnitEstDelDate) Then EstDateX = "" Else EstDateX = Me.UnitEstDelDate
        If IsNull(Me.UnitPDStartDate) Then PDateX = "" Else PDateX = Me.UnitPDStartDate
        Me.ClientGroupID = DLookup("ClientGroupId", "ClientGroup_MatchedClients", "[ClientId]=" & Me.clientID & "")
        StatusX = "B"
        If Me.MLOrig <> "FA" Then StatusX = "T"
SQLString:
        SQLz = "INSERT INTO Units (UnitBuildDate, UnitEstDelDate, CofASigner, UnitFinalCost, UnitDeliveryLocationRef, UnitGarageLocationRef, UnitTitleLocationRef, " & _
            "UnitRegLocationRef, UnitIRP, UnitTaxExempt, UnitTaxNotes, UnitPDStartDate, UnitRent, UnitsRent, UnitPDRent, MTMRent, FirstExtRent, UnitTAX, UnitSalesResid, " & _
            "UnitYr, UnitMake, UnitModel, UnitType, UnitSubType, UnitDel, UnitDep, UnitDepdt, Axle "

        If Nz(Me.UnitUOM, 0) <> 0 Then SQLz = SQLz & ",UnitUOM "
        If Nz(Me.UnitConfiguration, "") <> "" Then SQLz = SQLz & ",UnitConfiguration "
        If Nz(Me.UnitAxleProperty, "") <> "" Then SQLz = SQLz & ",UnitAxleProperty "
        If Nz(Me.UnitDescriptor, "") <> "" Then SQLz = SQLz & ",UnitDescriptor "

        SQLz = SQLz & ", GroupRef"
        SQLz = SQLz & ", VendorRef"
        SQLz = SQLz & ", ClientGroupID"
        SQLz = SQLz & ", UnitStatus"
        SQLz = SQLz & ", ClientID ) "
        SQLz = SQLz & "SELECT "
        SQLz = SQLz & "'" & BuildDateX & "'"
        SQLz = SQLz & ", '" & EstDateX & "'"
        SQLz = SQLz & ", " & Nz(Me.CofASigner, 0) & ""

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
        'SQLz = SQLz & ", '" & Me.MaintProv & "'"
        SQLz = SQLz & ", '" & Me.UnitDel & "'"
        SQLz = SQLz & ", " & Me.UnitDep & ""
        SQLz = SQLz & ", #" & Nz(Me.UnitDepdt, #1/1/2050#) & "#"
        SQLz = SQLz & ", '" & Me.axle & "'"
        If Nz(Me.UnitUOM, 0) <> 0 Then SQLz = SQLz & "," & Me.UnitUOM & " "
        If Nz(Me.UnitConfiguration, "") <> "" Then SQLz = SQLz & ",'" & Me.UnitConfiguration & "'"
        If Nz(Me.UnitAxleProperty, "") <> "" Then SQLz = SQLz & ",'" & Me.UnitAxleProperty & "'"
        If Nz(Me.UnitDescriptor, "") <> "" Then SQLz = SQLz & ",'" & Replace(Me.UnitDescriptor, "'", "''", 1) & "'"
        SQLz = SQLz & "," & Me.GroupRef & ", " & Me.VendorRef & ", " & Me.ClientGroupID & ", '" & StatusX & "', " & Me.clientID & ";"
        MessageUser "Open", "Adding Units..", "Time Flys"
        DoCmd.SetWarnings False
        For i = 1 To NumOfUnits
            MessageUser "Update", "Adding Unit " & i & " of: " & NumOfUnits, "Working it..."
            DoCmd.RunSQL (SQLz)
        Next i
        MessageUser "Update", "Doing Some Data Houskeeping", NumOfUnits & " Units, this will take a bit..."
        SQLz = "Update Units Set UnitUnitNum = 'TBD-' & UnitID " & _
            ",UnitVIN ='TBD-' & UnitId " & _
            ",ObcUnitNumber = 'TBD-' & UnitId  " & _
            "WHERE UnitVIN IS NULL and GroupREf=" & Me.GroupRef
        DoCmd.RunSQL (SQLz)
        MessageUser "Close"
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

Private Sub OpenUnitsGroup_Click()
    DoCmd.OpenForm "GroupForm", , , "SGrpID = " & Me.GroupRef
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
    '  If Me.unitstatus = "T" Or Me.unitstatus = "S" Or Me.unitstatus = "R" Or Me.unitstatus = "PS" Then
    '      Response = MsgBox("You are attempting to add a Log for a Unit that may not belong to " & vbNewLine & "Fleet Advantage. This Unit may be a Purchase Lease Back or in Remarketing" & vbNewLine & "Do you want to proceed?", vbQuestion + vbYesNo, "Not a FA unit")
    ' Else
    '      Response = MsgBox("You're about to ADD a Title Log Record, Are you Sure?", vbQuestion + vbYesNo, "New Record")
    '  End If
    DoCmd.close acForm, "OPSTitleLogEntry"
    DoCmd.OpenForm "OPSTitleLogEntry", , , , acFormAdd
    Forms!OPSTitleLogEntry.UnitREF = Me.UnitID
    Forms!OPSTitleLogEntry.USERID = FAUserName
    Forms!OPSTitleLogEntry.TitleNumber = Me.UnitTitleNumber
    If IsNull(Me.UnitRegLocationRef) Or Me.UnitRegLocationRef = 0 Then
        On Error Resume Next
        StateX = Nz(Forms!Unitsform!UNittitleLogForm.Form.TitleState, "")
CheckStatus:
        On Error GoTo 0
    Else
        StateX = Nz(DLookup("LState", "Locations", "LocationID =" & Me.UnitRegLocationRef), "")
    End If
    If Len(StateX) > 1 Then Forms!OPSTitleLogEntry.TitleState = StateX
    Exit Sub

End Sub

Private Sub SaveMe_Click()
    If Me.Form.Dirty = True Then Me.Form.Dirty = False
    DoCmd.Save
End Sub

Private Sub UnitActualDelDate_AfterUpdate()
    If Me.MLOrig = "FA" Then Call RePairExchangeIT
End Sub

Private Sub UnitCOABack_AfterUpdate()
    If Me.MLOrig = "FA" Then
        If IsDate(Me.UnitCOABack) And (Me.unitstatus) = "B" Then Me.unitstatus.Value = "A"
        If Not IsDate(Me.UnitCOABack) And Not IsDate(Me.UnitCofACopy) And (Me.unitstatus) = "A" Then Me.unitstatus.Value = "B"
    End If
End Sub

Private Sub UnitCofACopy_AfterUpdate()
    If Me.MLOrig = "FA" Then
        If IsDate(Me.UnitCofACopy) And (Me.unitstatus) = "B" Then Me.unitstatus.Value = "A"
        If Not IsDate(Me.UnitCOABack) And Not IsDate(Me.UnitCofACopy) And (Me.unitstatus) = "A" Then Me.unitstatus.Value = "B"
    End If
End Sub
Private Sub UnitAcceptDate_AfterUpdate()
    If Me.MLOrig = "FA" Then
        If IsDate(Me.UnitCofACopy) And Me.unitstatus = "B" Then Me.unitstatus.Value = "A"
        If IsDate(Me.UnitCOABack) And Me.unitstatus = "B" Then Me.unitstatus.Value = "A"
        Call RePairExchangeIT
    End If
End Sub

Private Sub UnitFinalCost_AfterUpdate()
    Call MakePaymentAdjustment(Me.GroupRef)
End Sub

Private Sub UnitGarageLocationRef_DblClick(Cancel As Integer)
    DoCmd.close acForm, "LocationAddressesEdit"
    DoCmd.OpenForm "LocationAddressesEdit", acNormal, , , , , DLookup("ClientGroupID", "Locations", "LocationID=" & UnitGarageLocationRef & "") & ":" & UnitGarageLocationRef
End Sub

Private Sub UnitIFAmt_AfterUpdate()
    If Me.UnitIFAmt > Me.VendorCost Then
        Call MsgBox("Sorry, you cannot enter a Value greater then the Vendor Cost." & vbNewLine & "This will be undone", vbCritical, "Stop spending money you don't have!")
        Me.Undo
    End If
End Sub

Private Sub UnitInServ_DblClick(Cancel As Integer)
    Me.UnitInServ.Locked = False
    If Me.MLOrig = "FA" Then Call RePairExchangeIT
End Sub

Private Sub UnitMake_Click()
    Me.UnitModel.Requery
End Sub
Public Sub unitoffleasedt_GotFocus()
    PrevOffLease = Nz(Me.unitoffleasedt, "1/1/1900")
End Sub

Private Sub unitoffleasedt_AfterUpdate()

    If DCount("UnitRef", "RMKT", "UnitREF =" & Me.UnitID) = 0 Then
        Call InsertRMKTUnit(Me.UnitID)
        Me.Form.Units_RMKT_Sub.Requery
    End If

    If IsDate(Me.Units_RMKT_Sub!parkdt) Then
        Select Case UCase(Me.unitstatus)
            Case "S", "D", "C", "N"   ' Do nothing
            Case Else
                Me.unitstatus = "R"
        End Select
    End If

    If PrevOffLease <> Me.unitoffleasedt Then
        Dim PreDatex As String
        If PrevOffLease = "1/1/1900" Then PreDatex = "Null" Else PreDatex = PrevOffLease
        Committ = MsgBox("You're about to Change the Off Lease Date: " & vbNewLine & " From: " & PreDatex & "  to: " & unitoffleasedt & vbNewLine & "Do You Agree?", vbQuestion + vbYesNo + vbCritical, "Change Offlease?")
        If Committ <> 6 Then
            If PrevOffLease = "1/1/1900" Then Me.unitoffleasedt = Null Else Me.unitoffleasedt = PrevOffLease
            Exit Sub
        End If
    End If
    If DLookup("LeaseType", "SchGrp", "sGrpid=" & Me.GroupRef) = "SLB Insp" Then
        If Me.Dirty = True Then Me.Dirty = False
        DoCmd.SetWarnings False
        Me.UnitsTabControl.Value = 6
        'Call SetSubData(6)
        If Me.unitoffleasedt & "test" = "test" Then Call SetSlbInspBuyOut(Me.UnitID, Me.GroupRef, True) Else Call SetSlbInspBuyOut(Me.UnitID, Me.GroupRef, False)
        Me.Form.Units_RMKT_Sub.Requery
    End If
    DoCmd.SetWarnings True
End Sub
Public Sub SetSlbInspBuyOut(AssetIDx As Long, GroupIDx As Long, ClearOffLease As Boolean)
    If DLookup("LeaseType", "SchGrp", "sGrpid=" & GroupIDx) <> "SLB Insp" Then Exit Sub
    Dim BLCDx As Date
    Dim OffLeaseDateX As Date
    Dim PdRent As Currency
    Dim VCost As Currency
    Dim PrevBuyOut As Currency
    Dim Committ As String
    Dim RST As DAO.Recordset
    If ClearOffLease = True Then
        SQLz = "Update RMKT Set BuyOutCost=0 , BuyOutDate=Null Where UnitRef=" & AssetIDx & ";"
        GoTo UpdateRMKT
    End If
    SQLz = "Select * From vw_SixKeys Where AssetID=" & AssetIDx
    Set RST = CurrentDb.OpenRecordset(SQLz)
    BLCDx = RST!BLCD
    OffLeaseDateX = Nz(RST!OffLeaseActual, Date)
    VCost = RST!VendorCost + Nz(PartsCost, 0)
    PdRent = DLookup("PerDiemSurrender", "Units", "UnitID=" & AssetIDx)
    RentIncome = (DateDiff("d", BLCDx, OffLeaseDateX) + 1) * PdRent
    PrevBuyOut = Nz(DLookup("BuyOutCost", "RMKT", "UnitREf=" & AssetIDx), 0)
    RST.close
    Set RST = Nothing
    If PrevBuyOut > 0 Then
        MsgBox "This Changed the BuyOut Cost" & vbNewLine & "From: " & Format(PrevBuyOut, "$#,#.##") & "  to: " & Format(VCost - RentIncome, "$#,#.##") & "", vbInformation, "Buyout Changed!"
    End If
    SQLz = "Update RMKT Set BuyOutCost=" & (VCost - RentIncome) & " , BuyOutDate='" & OffLeaseDateX & "' Where UnitRef=" & AssetIDx & ";"
UpdateRMKT:
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
End Sub

Private Sub UnitsTabControl_Change()
    Call SetSubData(Me.UnitsTabControl.Value)
End Sub
Public Function SetSubData(TabID As Integer)
    If TabID = 0 Then    ' Locations Tab
        Call Me.UnitLocationChangeHistoryForm.Form.LoadData(Me.UnitID)
        Call Me.UNittitleLogForm.Form.LoadData(Me.UnitID)
    End If
    If TabID = 1 Then    ' Unit Notes Tab
        Call Me.UnitNotes.Form.LoadData(Me.UnitID)
    End If
    If TabID = 2 Then    ' Child Parts Tab
        Call Me.ChildParts_sub.Form.LoadData(Me.UnitID)
    End If
    If TabID = 3 Then    ' Vin Decoder Tab
        Call Me.VINDecode.Form.LoadData(Me.UnitID)
    End If
    If TabID = 4 Then    ' Mile/Hours Tab
        Call Me.ViewMiNew.Form.LoadData(Me.UnitID)
    End If
    If TabID = 5 Then    ' Recalls Tab
        Call UnitRcallsub.Form.LoadData(Me.UnitID)
    End If
    If TabID = 6 Then    ' Remarketing Tab
        Me.SUFHistoryList.Requery
        Call Me.Units_RMKT_Sub.Form.LoadData(Me.UnitID)
    End If
    If TabID = 7 Then    ' InspWorkOrder Tab
        Call Me.InspectionsSubForm.Form.LoadData(Me.UnitID)
        Me.HrParked.Value = Nz(DLookup("HrParked", "RMKT", "UnitRef=" & Me.UnitID & ""), 0)
        Me.MiParked.Value = Nz(DLookup("MiParked", "RMKT", "UnitRef=" & Me.UnitID & ""), 0)
    End If
    If TabID = 8 Then    ' Exit Options Tab
        Call Me.Units_Exit_ExchangeIT.Form.LoadData(Me.UnitID)
        Call Me.Units_Exit_Extend.Form.LoadData(Me.UnitID)
        Call Me.Units_Exit_Terminate.Form.LoadData(Me.UnitID)
    End If
    If TabID = 9 Then    ' UnitChanges Tab
        Call Me.UnitRevisionLog.Form.LoadData(Me.UnitID)
    End If
End Function

Private Sub UnitStatus_Click()
    If Me.unitstatus = "A" Then
        If IsNull(Me.UnitCOABack) And IsNull(UnitCofACopy) Then
            MsgBox "Thus unit cannot be set to 'A' without a COA-Back or a COA-Copy date", vbCritical, "Not yet an Accepted Unit"
            Me.Form.Undo
        End If
    End If
End Sub

Private Sub UnitStatus_GotFocus()
    PrevStatus = Me.unitstatus
    If FAUserAccessLevel >= 6 Or FAUserAccessGroup Like "*OPS*" Then
        Me.unitstatus.Locked = False
    Else
        If Me.unitstatus = "R" Or Me.unitstatus = "S" Or Me.unitstatus = "D" Then
            Me.unitstatus.Locked = True
        End If
    End If
End Sub

Private Sub UnitSubType_Click()
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
        SQLz = SQLz & "SELECT " & Me.UnitID & ", 'Title Application Sent', #" & Me.UnitTASent & "#, 2, 18, True, '" & FAUserName & "';"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
        Me.UNittitleLogForm.Requery
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

    If Me.unitstatus <> "A" Or Me.unitstatus <> "B" Then
        DoCmd.close acForm, "UnitsForm", acSaveYes
        Exit Sub
    End If
    Dim Msgx   As String
    Msgx = "| "
    If IsNull(Me.UnitYr.Value) Then Msgx = Msgx & "UnitYr Missing |"
    If IsNull(Me.UnitMake) Then Msgx = Msgx & "UnitMake Missing |"
    If IsNull(Me.UnitType) Then Msgx = Msgx & "UnitType Missing |"
    If IsNull(Me.VendorRef) Then Msgx = Msgx & "Vendor Missing |"
    If unitstatus = "A" Or unitstatus = "B" Then
        If Me.UnitFinalCost.Value <= 0 Then Msgx = Msgx & "Final Cost not set |"
    End If
    If Msgx <> "| " Then
        Me.Undo
        MsgBox "Required fields are missing" & vbNewLine & Msgx & vbNewLine & "No Changes Made!", vbCritical, "Required Fields Missing"
        DoCmd.close acForm, Me.Form.Name, acSaveNo
        Exit Sub
    End If
    If (IsNull(Me.UnitGarageLocationRef) Or (Me.UnitGarageLocationRef) = 0) And (Me.unitstatus) = "A" Then
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
        UCCClogDet("User").Value = FAUserName
        UCCClogDet.Update

    End If
End Sub

Private Sub UnitGarageLocationRef_GotFocus()
    If FAUserAccess < 16 Then Me.UnitGarageLocationRef.Locked = True Else Me.UnitGarageLocationRef.Locked = False
    OldLocationLng = Nz(Me.UnitGarageLocationRef, 0)
End Sub

Private Sub UnitGarageLocationRef_AfterUpdate()
    If OldLocationLng <> 0 And OldLocationLng <> Me.UnitGarageLocationRef.Value Then
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
            AsOfDatex = InputBox("What is the Actual Date the Unit Moved?" & vbNewLine & "please keep correct date format mm/dd/yyyy", "Date this Unit Actually moved.", Format(Date, "mm/dd/yyyy"))
            If Not IsDate(AsOfDatex) Then GoTo InputDate
            'On Error GoTo FixMe
            SQLz = "INSERT INTO LocationChangeHistory ( AssetId, OldLocationId, NewLocationId, LocationChangeDate, LocationType, Reason, GroupId, FAUserName ) " & _
                "SELECT " & Me.UnitID & " AS AssetId, " & OldLocationLng & " AS OldLocationId, " & Me.UnitGarageLocationRef.Value & " AS NewLocationId, #" & AsOfDatex & "# AS LocationChangeDate, 'Garage'   AS LocationType, '" & ReasonX & "' as Reason, " & Me.GroupRef & "  AS GroupId, FAUserName();"
            DoCmd.SetWarnings False
            DoCmd.RunSQL (SQLz)
            DoCmd.SetWarnings True
            Me.UnitLocationChangeHistoryForm.Form.Requery
        Else
            Me.UnitGarageLocationRef = OldLocationLng
            Me.UnitGarageLocationRef.Requery
            OldLocationLng = Nz(Me.UnitGarageLocationRef, 0)
        End If
        OldLocationLng = Me.UnitGarageLocationRef
    End If
    Call SyncUnitsToLocationDivisions
FixMe:
End Sub

Private Sub UnitGarageLocationRef_Dirty(Cancel As Integer)
    Exit Sub
End Sub

Private Sub UnitStatus_BeforeUpdate(Cancel As Integer)
    If Me.unitstatus & "zz" = "zz" Then
        MsgBox "Unit Status cannot be blank!", vbCritical, "Error!"
        Me.unitstatus.Undo
        Exit Sub
    End If
    If Me.UnitGarageLocationRef & "zz" = "zz" Then
        MsgBox "Please add Garage Location and then update the Unit Status"
        Cancel = True
        Me.unitstatus.Undo
        Exit Sub
    End If
    If PrevStatus = "C" Then
        If Me.unitstatus <> "C" Then
            MsgBox "Once a Unit is a Casualty ('C'), it will remain a 'C' Status Unit, even when Parked or Off Leased.", vbCritical, "Just Stop"
        End If
        Me.unitstatus = "C"
    End If
    If Me.MLOrig <> "FA" Then
        If Me.unitstatus <> "T" Or Me.unitstatus <> "O" Then
            MsgBox "Only FA Units can have Status: '" & Me.unitstatus & "'", vbCritical, "Are you looking for trouble?"
            Cancel = True
            Me.unitstatus.Undo
            Exit Sub
        End If
    Else
        If Me.unitstatus = "T" _
            Or Me.unitstatus = "Ti" _
            Or Me.unitstatus = "O" Then
            MsgBox "Only Tracked Units can have Status: '" & Me.unitstatus & "'", vbCritical, "Are you looking for trouble?"
            Cancel = True
            Me.unitstatus.Undo
            Exit Sub
        End If
    End If

    If Me.unitstatus = "'Ti" Or Me.unitstatus = "SB" Or Me.unitstatus = "O" Then
        If Not IsDate(Me.unitoffleasedt) Then
            MsgBox "To Off-Lease a Unit, You must have an Offlease Date Entered.", vbCritical, "No Off Lease Date!"
            Me.unitstatus.Undo
        End If
    End If
    If Me.unitstatus = "P" Or Me.unitstatus Like "*S*" Or Me.unitstatus = "D" Or Me.unitstatus = "'AR" Or Me.unitstatus = "Ti" Or Me.unitstatus = "TR" Then
        If IsUnitInRMKT(Me.UnitID) < 1 Then
            Call InsertRMKTUnit(Me.UnitID)
            Call Units_RMKT_Sub.Form.LoadData(Me.UnitID)
        End If
        If IsNull(Me.Units_RMKT_Sub!parkdt) And Me.unitstatus = "P" Then
            MsgBox "You must have a Parked Date set before you can Park ('P') a unit"
            Cancel = True
            Me.unitstatus.Undo
        End If
    End If
End Sub

Private Sub UnitStatus_DblClick(Cancel As Integer)
    If Me.AllowAdditions = True And Me.AllowEdits = True And Me.AllowDeletions = True Then
        Exit Sub
    Else
        DoCmd.OpenForm "unlock unit form"
    End If
End Sub



Private Sub UnitType_Click()
    Me.UnitSubType.Requery
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
