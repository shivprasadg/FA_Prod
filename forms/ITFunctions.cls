Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Dim SQLz       As String



Private Sub ATFStructure_Click()
    DoCmd.close acForm, "ATF_Structure"
    DoCmd.OpenForm "ATF_Structure"
End Sub

Private Sub CloseMe_Click()
    DoCmd.OpenForm "FAMainMenu"
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub CostcoImporter_Click()
    Call OpenMnRDataFileDialog("Costco", Me.Date2Process)
End Sub

Private Sub DecodeFromList_Click()
Call OpenGetVINDialog
End Sub

Private Sub DecodeVINS_Click()
    If Me.DecodeClient & "test" = "test" Then Exit Sub
    Dim RST    As DAO.Recordset
    Dim xCount, i, x As Integer

    Call MessageUser("Open", "Getting VIN Data", "working...")

    SQLz = "SELECT S.clientID, S.UnitID, S.VIN, S.MYear, S.Type " & _
        "FROM vw_SixKeys as S LEFT JOIN UnitVinDetails AS D ON S.UnitID = D.unitid "
    SQLz = SQLz & "WHERE S.clientID=" & Me.DecodeClient.Column(1) & " "
    SQLz = SQLz & "AND D.UnitId IS NULL "
    SQLz = SQLz & "AND (S.Type ='" & Me.DecodeClient.Column(6) & "') "
    '--SQLz = SQLz & "AND (S.Type ='Tractor' Or S.Type='Straight Truck' Or S.Type='Van') "
    SQLz = SQLz & "AND Len(S.VIN) = 17 "
    SQLz = SQLz & "AND Not S.VIN Like '% %' "
    SQLz = SQLz & "AND (S.Status='" & Me.DecodeClient.Column(4) & "');"
    'SQLz = SQLz & "AND (D.Status Is Null or D.Status='" & Me.DecodeClient.Column(5) & "') "
    On erorr GoTo FixMe

    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    If RST.RecordCount > 0 Then
        RST.MoveLast
        RST.MoveFirst
        If Me.oDecodeOne = True Then xCount = 1 Else xCount = RST.RecordCount
        i = 1
        For i = 1 To xCount
            If Me.oDecodeOne = True Then
                Call MessageUser("Update", "Clone VINGroup: " & Left(RST!VIN, 8) & Mid(RST!VIN, 10, 1), _
                    "Attempt Unit " & i & " Of " & RST.RecordCount)
            Else
                Call MessageUser("Update", "VIN: " & RST!VIN, "Unit " & i & " Of " & xCount)
            End If
            DoCmd.SetWarnings False
            DoCmd.RunSQL ("Delete From UnitVinDetails WHERE UnitID =" & RST!UnitID)
            DoCmd.SetWarnings True
            If oUseMYear = True Then
                Call VinDecoder(RST!UnitID, Nz(RST!VIN, "NoVIN"), Nz(RST!MYear, ""), True, False, RST!ClientID)
            Else
                Call VinDecoder(RST!UnitID, Nz(RST!VIN, "NoVIN"), , True, False, RST!ClientID)
            End If
            RST.MoveNext
        Next i
        If Me.oDecodeOne = True Then
            x = CloneDecodedVINS
            Call MessageUser("Update", " VINGroup VIN Cloned:", x & " Units")
            ' If x > 0 Then Pause (0)
        End If
    End If
    Me.DecodeClient.Requery
    If Nz(Me.DecodeClient.Column(6), "") <> "Trailer" Then Call SetEngineType(True)
    If Me.oDecodeOne = True And Repeatclone = True And Me.DecodeClient.Column(1) & "test" <> "test" Then Call DecodeVINS_Click
FixMe:
    Call MessageUser("Close")
    On Error GoTo 0
End Sub
Private Sub CheckVINsOnly_Click()
'If Me.DecodeClient & "test" = "test" Then Exit Sub
    Dim RST    As DAO.Recordset
    Dim xCount, i As Integer
    SQLz = "SELECT UnitID, ClientID, UnitYear, VIN FROM Unit_VIN_Temp;"

    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)

    If RST.RecordCount > 0 Then
        RST.MoveLast
        RST.MoveFirst
        xCount = RST.RecordCount
        i = 1
        Call MessageUser("Open", "Getting VIN Data", "working...")
        Do Until RST.EOF
            Call MessageUser("Update", "VIN: " & RST!VIN, "Unit " & i & " Of " & xCount)
            If Nz(RST!UnitYear, 0) > 0 Then
                Call VinDecoder(Nz(RST!UnitID, 0), Nz(RST!VIN, "NoVIN"), RST!UnitYear, True, True, Nz(RST!ClientID, 0))
            Else
                Call VinDecoder(Nz(RST!UnitID, 0), Nz(RST!VIN, "NoVIN"), , True, True, Nz(RST!ClientID, 0))
            End If
            RST.MoveNext
            i = i + 1
        Loop
        Call MessageUser("Close")
    End If
    Me.DecodeClient.Requery
    Call SetEngineType(True)
End Sub

Private Sub FixEngineMFG_Click()
    Call SetEngineType(False)
End Sub

Private Sub GetUnDecodedVINS_Click()

MessageUser "Open", "Getting All Not Decoded VINs", "Inserting them into Temp Table."
SQLz = "Delete From Unit_Vin_Temp "
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
SQLz = "Insert into Unit_Vin_Temp (ClientID,UnitID,VIN, UnitYear) "
SQLz = SQLz & "SELECT ClientID, UnitID,UnitVIN,UnitYr "
SQLz = SQLz & "FROM Units "
SQLz = SQLz & "WHERE UnitEngineType Is Null "
SQLz = SQLz & "AND UnitType<> 'Trailer' "
SQLz = SQLz & "AND Len(UnitVin) = '17' ORDER BY ClientID;"
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
MessageUser "Update", "Calling Decoder", "This will take a moment or two..."
Call DecodeVinList("", , True, True, True)
End Sub

Private Sub ImportFAIDFromSalesForce_Click()
    DoCmd.OpenForm "ATF_SFImporter"
End Sub

Private Sub LinkReleasedUnits_Click()
    Call MessageUser("Open", "Linking Units", "Temp table delete")
    On Error Resume Next
    DoCmd.SetWarnings False
    DoCmd.DeleteObject acTable, "ReleasedTemp"
    DoCmd.SetWarnings True
    On Error GoTo 0
    SQLz = ""
    SQLz = SQLz & "SELECT vw_ReLeasedUnits.NewAssetID, vw_ReLeasedUnits.AssetID_FATitleLog "
    SQLz = SQLz & "INTO ReleasedTemp FROM vw_ReLeasedUnits;"
    Call MessageUser("update", "Linking Units", "Making Temp table")
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    SQLz = ""
    SQLz = SQLz & "UPDATE ReleasedTemp "
    SQLz = SQLz & "INNER JOIN Units ON ReleasedTemp.NewAssetID = Units.UnitID "
    SQLz = SQLz & "SET Units.AssetID_FATitleLog = ReleasedTemp.AssetID_FATitleLog "
    SQLz = SQLz & "WHERE Units.UnitID=units.AssetID_FATitleLog;"
    Call MessageUser("update", "Linking Units", "Making Update to units table")
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.DeleteObject acTable, "ReleasedTemp"
    DoCmd.SetWarnings True
    Call MessageUser("Close")
End Sub

Private Sub PenskeImporter_Click()
    Call OpenMnRDataFileDialog("Penske", Me.Date2Process)
End Sub

Private Sub Fleet_Importer_Click()
    DoCmd.close acForm, "FleetStudy_Importer"
    DoCmd.OpenForm "FleetStudy_Importer"
End Sub
Private Sub GeoMap_Click()
    DoCmd.OpenForm "MapPopUp1"
End Sub
Private Sub MakeManager_Click()
    DoCmd.OpenForm "MakeModelManager"
End Sub

Private Sub OBCUnitnumberCheck_Click()
    Call MessageUser("Open", "OBC Unit Number Update", "This may take a moment, or two")
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "UnitUnitNun_2_OBCUnitNumber"
    Call MessageUser("Update", "OBC Unit Number Check Process Complete", "almost done...")
    DoCmd.RunSQL ("UPDATE Units SET Units.MpgGoal = 6.5, Units.DrivingMpgGoal = 6.5 WHERE (((Units.MpgGoal)<1 Or (Units.MpgGoal) Is Null));")
    Call MessageUser("Update", "MPG Goals set", "done")
    DoCmd.SetWarnings True
    Call MessageUser("Close")
End Sub
Private Sub MoveUnitsForm_Click()
    DoCmd.OpenForm "MoveUnit"
End Sub

Private Sub PortalReports_Click()
'Sets Report-Start-End Date for portal reports. If outside of date range the units will not show up in reports.
    Dim RST    As DAO.Recordset
    Dim xCount, Ycount As Integer
    xCount = 0
    Ycount = 0
    Call MessageUser("Open", "Setting Report Start Dates", "...Collecting Data...")
    DoCmd.SetWarnings False
    SQLz = "Select * From UnitReportStartEndDates_Start Order By UnitId;"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    If RST.RecordCount > 0 Then
        RST.MoveLast
        RST.MoveFirst
        xCount = RST.RecordCount
        Do Until RST.EOF
            SQLz = "UPDATE Units SET Units.ReportStartDate = #" & RST!StartDate & "#, Units.ReportEndDate = #" & RST!EndDate & "# " & _
                   "WHERE Units.UnitID = " & RST!UnitID & ";"
            Call MessageUser("Update", "Posting Start Dates", "Setting data for AssetID:" & RST!UnitID)
            DoCmd.RunSQL (SQLz)
            RST.MoveNext
        Loop
    End If
    SQLz = "Select * From UnitReportStartEndDates_End Order By UnitId;"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    If RST.RecordCount > 0 Then
        RST.MoveLast
        RST.MoveFirst
        Ycount = RST.RecordCount
        DoCmd.SetWarnings False
        Do Until RST.EOF
            SQLz = "UPDATE Units SET Units.ReportStartDate = #" & RST!StartDate & "#, Units.ReportEndDate = #" & RST!EndDate & "# " & _
                   "WHERE Units.UnitID = " & RST!UnitID & ";"
            Call MessageUser("Update", "Posting End Dates", "Setting data for AssetID:" & RST!UnitID)
            DoCmd.RunSQL (SQLz)
            RST.MoveNext
        Loop
    End If
    DoCmd.SetWarnings True
    Call MessageUser("Close")
    MsgBox "Reports Start-End Complete" & vbNewLine & xCount & " Start Dates Updated" & vbNewLine & Ycount & " End Dates Updated", , "Month End Process"
End Sub

Private Sub ReconnectATLAASCore_Click()
    Call ChangeConnectionString("FA-Phoenix", "FleetAdvantage_beSQL")
    Forms.FAMainMenu.DevLabel.Visible = True
End Sub

Private Sub ReconnectPoseidon_Click()
    'Call ChangeConnectionString("fa-poseidon.eastus.cloudapp.azure.com,1549", "FleetAdvantage_beSQL", , 0)
'    Call ChangeConnectionString("fa-poseidon.domain.local,1549", "FleetAdvantage_beSQL", , 0)
    Call ChangeConnectionString
    Forms.FAMainMenu.DevLabel.Visible = True
End Sub

Private Sub ReconnectPhoenix_Click()
    Call ChangeConnectionString("FA-Phoenix", , , False)
    'Call ChangeConnectionString("192.168.2.5", , , False)

    Forms.FAMainMenu.DevLabel.Visible = False
End Sub

Private Sub CompressATLAAS_Click()
    CommandBars("Menu Bar").Enabled = True
    CommandBars("Menu Bar").Controls("Tools").Controls("Database utilities").Controls("Compact and repair database").accDoDefaultAction
End Sub

Private Sub RunAutoExtend__Click()
  Call AutoExtendMTM(False)
End Sub

Private Sub SignOffTasks_Click()
  DoCmd.close acForm, "SignOff Tasks"
  DoCmd.OpenForm "SignOff Tasks"
End Sub

Private Sub SYGMAPenske_Click()
    Call MessageUser("Open", "Clearing Previous SYGMA MnR Data", "just a moment or two, click OK when asked.")
    SQLz = "Delete * From MandRData_BETL Where ClientID=3674"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Call MessageUser("Update", "Posting SYGMA Penske MnR Data...", "Working 2 to 3 mins, click OK when asked")
    SQLz = ""
    SQLz = SQLz & " INSERT INTO MandRData_BETL ( ClientGroupId, ClientID, UnitID, UnitNumber, VIN, FleetCodeDescr, CompanyCodeDescr"
    SQLz = SQLz & ", CompanyDivision, TypeDescr, VMRS, SystemCodeDescr, RoNumber, RODateTime, ActivityMonth, [Usage], LTD"
    SQLz = SQLz & ", LaborHours, LaborCost, PartCost, OutSideCost, TotalCost )"
    SQLz = SQLz & "SELECT"
    SQLz = SQLz & "  Units.ClientGroupId"
    SQLz = SQLz & ", Units.ClientID"
    SQLz = SQLz & ", Units.UnitID"
    SQLz = SQLz & ", MandRSygma_Penske.[Penske Unit #] AS UnitNumber"
    SQLz = SQLz & ", MandRSygma_Penske.[VIN #]"
    SQLz = SQLz & ", Propercase([Area],1) AS FleetCodeDescr"
    SQLz = SQLz & ", Propercase([Customer],1) AS CompanyCodeDescr"
    SQLz = SQLz & ", Propercase([Region],1) AS CompanyDivision"
    SQLz = SQLz & ", Units.UnitType As TypeDescr"
    SQLz = SQLz & ", MandRSygma_Penske.[Comp Code] AS VMRS"
    SQLz = SQLz & ", Propercase([Comp Code Desc],1) AS SystemCodeDescr"
    SQLz = SQLz & ", MandRSygma_Penske.[RO #] AS RoNumber"
    SQLz = SQLz & ", [RO Date] AS RODateTime"
    SQLz = SQLz & ", CDate(SplitString([RO Date],'/',0) & '/1/' & SplitString([Ro Date],'/',2)) AS ActivityMonth"
    SQLz = SQLz & ", MandRSygma_Penske.Usage"
    SQLz = SQLz & ", MandRSygma_Penske.Usage AS LTD"
    SQLz = SQLz & ", MandRSygma_Penske.[Labor Hrs] AS LaborHours"
    SQLz = SQLz & ", CCur(Nz([Labor Amt],0)) AS LaborCost"
    SQLz = SQLz & ", CCur(Nz([Parts Amt],0)) AS PartCost"
    SQLz = SQLz & ", CCur(Nz([Vendor Amt],0)) AS OutSideCost"
    SQLz = SQLz & ", CCur(Nz([Total Amt Invoiced],0)) AS TotalCost "
    SQLz = SQLz & "FROM MandRSygma_Penske INNER JOIN Units ON MandRSygma_Penske.[VIN #] = Units.UnitVIN "
    SQLz = SQLz & "WHERE (((MandRSygma_Penske.[VIN #])<>''));"
    DoCmd.RunSQL (SQLz)
    Call MessageUser("close")
    Call SendImportEmailMsg("Sygma-Penske", 88, DatePart("m", Date) - 1 & "/1/" & DatePart("yyyy", Date), DCount("ClientID", "MandRData_BETL", "ClientID=3674"))
End Sub

Private Sub SYGMAVector_Click()
    Call GetMnRData_Vector(, Date2Process)
End Sub

Private Sub SyncLocationDivisions_Click()
    Call SyncUnitsToLocationDivisions

    MsgBox "Locations / Divisions Synced to the Divisions Assigned to the Units' Locations" & vbNewLine & "(if any needed to change)", vbInformation
End Sub

Private Sub TranscoMnR_Click()
    Call MessageUser("Open", "Clearing Previous Transco MnR Data", "This will take a bit, click OK when asked")
    SQLz = "Delete * From MandRData_BETL Where ClientID=4971"
    DoCmd.RunSQL (SQLz)

   ' SQLz = "Delete MandRTransco.txtSystemCodeDescr From MandRTransco WHERE MandRTransco.txtSystemCodeDescr Like '*No *';"
   ' DoCmd.RunSQL (SQLz)

    Call MessageUser("Update", "Posting Current Transco MnR Data", "Working, 2 to 3 min, click OK when asked")
    SQLz = ""
    SQLz = SQLz & "INSERT INTO MandRData_BETL ( VIN, FleetCodeDescr, CompanyCodeDescr, ClientGroupID, UnitID"
    SQLz = SQLz & ", ClientID, CompanyDivision, UnitNumber, TypeDescr, VMRS, SystemCodeDescr, RONumber, RODateTime"
    SQLz = SQLz & ", ActivityMonth, [Usage], LTD, LaborHours, LaborCost, OutsideCost, FluidCost, PartCost, TotalCost ) "
    SQLz = SQLz & "SELECT "
    SQLz = SQLz & " Units.UnitVIN"
    SQLz = SQLz & ",txteqFleetCodeDescr"
    SQLz = SQLz & ",txteqCompanyCodeDescr"
    SQLz = SQLz & ",Units.ClientGroupID"
    SQLz = SQLz & ",Units.UnitID"
    SQLz = SQLz & ",Units.ClientID"
    SQLz = SQLz & ",txtEqFleetCodeDescr as Division"
    SQLz = SQLz & ",txtEqUnitCode"
    SQLz = SQLz & ",Units.UnitType As TypeDescr"
    SQLz = SQLz & ",iif(txtSystemCodeDescr like '*No *','',Trim(SplitString(txtSystemCodeDescr, '-', 0))) as VMRS"
    SQLz = SQLz & ",txtSystemCodeDescr"
    SQLz = SQLz & ",txtRONumber"
    SQLz = SQLz & ",txtRODateTime"
    SQLz = SQLz & ",ActivityMonth"
    SQLz = SQLz & ",txtUsage"
    SQLz = SQLz & ",txtLTD"
    SQLz = SQLz & ",txtLaborHours"
    SQLz = SQLz & ",txtLaborCost"
    SQLz = SQLz & ",txtOutsideCost"
    SQLz = SQLz & ",txtFluidCost"
    SQLz = SQLz & ",txtPartCost"
    SQLz = SQLz & ",txtTotalCost "
    SQLz = SQLz & "FROM MandRTransco INNER JOIN Units ON MandRTransco.txtEqUnitCode = Units.UnitUnitNum "
    SQLz = SQLz & "WHERE MandRTransco.TxtEqCompanyCodeDescr <>'' And Units.ClientId=4971;"
    DoCmd.RunSQL (SQLz)
    MessageUser ("Close")
End Sub
Private Sub AirGasMnR_Click()
    Call MessageUser("Open", "Clearing Previous AirGas MnR Data", "This will take a bit, click OK when asked")
    SQLz = "Delete * From MandRData_BETL Where ClientID=3319"
    DoCmd.RunSQL (SQLz)
    Call MessageUser("Update", "Posting Current AirGas MnR Data", "Working, 2 to 3 min, click OK when asked")
    SQLz = ""
    SQLz = SQLz & "INSERT INTO MandRData_BETL ( VIN, FleetCodeDescr, CompanyCodeDescr, ClientGroupID, UnitID"
    SQLz = SQLz & ", ClientID, CompanyDivision, UnitNumber, TypeDescr, VMRS, SystemCodeDescr, RONumber, RODateTime"
    SQLz = SQLz & ", ActivityMonth, [Usage], LTD, LaborHours, LaborCost, OutsideCost, FluidCost, PartCost, TotalCost ) "
    SQLz = SQLz & "SELECT "
    SQLz = SQLz & " Trim(SERIALNO) as VIN "
    SQLz = SQLz & ",[Repair Shop Name] As FleetCodeDescr"
    SQLz = SQLz & ",'Airgas USA, LLC' As CompanyCodeDescr"
    SQLz = SQLz & ",Units.ClientGroupID"
    SQLz = SQLz & ",Units.UnitID"
    SQLz = SQLz & ",Units.ClientID"
    SQLz = SQLz & ",[Repair Shop] As CompanyDivision"
    SQLz = SQLz & ",UNITNUMBER "
    SQLz = SQLz & ",Units.UnitType As TypeDescr"
    SQLz = SQLz & ",CStr(VRMS)"
    SQLz = SQLz & ",[Comp Code Desc] As SystemCodeDescr"
    SQLz = SQLz & ",ORDERNUM As RONumber"
    SQLz = SQLz & ",Format([ClosedON],'mm/dd/yyyy') As RODateTime"
    SQLz = SQLz & ",Format(ActivityMonth,'mm/dd/yyyy')"
    SQLz = SQLz & ",0 As Usage"
    SQLz = SQLz & ",0 As LTD"
    SQLz = SQLz & ",0 As LaborHours"
    SQLz = SQLz & ",Labor As LaborCost"
    SQLz = SQLz & ",0 As OutsideCost"
    SQLz = SQLz & ",0 As FluidCost"
    SQLz = SQLz & ",Parts As PartCost"
    SQLz = SQLz & ",Linetotal as TotalCost "
    SQLz = SQLz & "FROM MandRAirGas INNER JOIN Units ON MandRAirGas.SERIALNO = Units.UnitVin "
    SQLz = SQLz & "WHERE MandRAirGas.UNITNUMBER <>'' And Units.ClientId=3319;"
    DoCmd.RunSQL (SQLz)
    MessageUser ("Close")

End Sub
Private Sub OldCastleMnR_Click()
    Call MessageUser("Open", "Clearing Previous OldCastle MnR Data", "This will take a bit, click OK when asked")
    SQLz = "Delete * From MandRData_BETL Where ClientID=4708"
    DoCmd.RunSQL (SQLz)
    Call MessageUser("Update", "Posting Current OldCastle MnR Data", "Working, 2 to 3 min, click OK when asked")
    SQLz = ""
    SQLz = SQLz & "INSERT INTO MandRData_BETL ( VIN, FleetCodeDescr, CompanyCodeDescr, ClientGroupID, UnitID"
    SQLz = SQLz & ", ClientID, CompanyDivision, UnitNumber, TypeDescr, VMRS, SystemCodeDescr, RONumber, RODateTime"
    SQLz = SQLz & ", ActivityMonth, [Usage], LTD, LaborHours, LaborCost, OutsideCost, FluidCost, PartCost, TotalCost ) "
    SQLz = SQLz & "SELECT "
    SQLz = SQLz & " Units.UnitVin as VIN "
    SQLz = SQLz & ",Location As FleetCodeDescr"
    SQLz = SQLz & ",'OldcastleBE' As CompanyCodeDescr"
    SQLz = SQLz & ",Units.ClientGroupID"
    SQLz = SQLz & ",Units.UnitID"
    SQLz = SQLz & ",Units.ClientID"
    SQLz = SQLz & ",[Location Code] As CompanyDivision"
    SQLz = SQLz & ",UNIT "
    SQLz = SQLz & ",Units.UnitType As TypeDescr"
    SQLz = SQLz & ",CStr(VRMS)"
    SQLz = SQLz & ",Description As SystemCodeDescr"
    SQLz = SQLz & ",[BD PO] As RONumber"
    SQLz = SQLz & ",Format([Completed Date],'mm/dd/yyyy') As RODateTime"
    SQLz = SQLz & ",Format(ActivityMonth,'mm/dd/yyyy')"
    SQLz = SQLz & ",0 As Usage"
    SQLz = SQLz & ",0 As LTD"
    SQLz = SQLz & ",0 As LaborHours"
    SQLz = SQLz & ",Labor + [service call] As LaborCost"
    SQLz = SQLz & ",[MISC] + [Wrecker] As OutsideCost"
    SQLz = SQLz & ",0 As FluidCost"
    SQLz = SQLz & ",[Parts]+[tires] + [Tire Disposal] As PartCost"
    SQLz = SQLz & ",[Total Cost] as TotalCost "
    SQLz = SQLz & "FROM MandROldcastle INNER JOIN Units ON MandROldcastle.Unit = Units.UnitunitNum "
    SQLz = SQLz & "WHERE MandROldcastle.UNIT <>'' And Units.ClientId=4708;"
    DoCmd.RunSQL (SQLz)
    MessageUser ("Close")

End Sub
Private Sub TrailerDataImporter_Click()
'Replaced by other routines
    Exit Sub
    SQLz = "INSERT INTO TrailerMilesHours ( LogActivityDate, UnitID, UnitNumber, Miles, MilesLTD, Hours, HoursLTD, ObcDataStatus, OBCDataDesc ) " & _
           "SELECT MainesTrailerData.LogActivityDate, Units.UnitID, Units.UnitUnitNum, MainesTrailerData.MilesDelta, MainesTrailerData.MILES, " & _
           "MainesTrailerData.HoursDelta, MainesTrailerData.RFRHR, MainesTrailerData.ObcDataStatus, MainesTrailerData.ObcDataStatusDesc " & _
           "FROM MainesTrailerData INNER JOIN Units ON MainesTrailerData.VIN = Units.UnitVIN;"
    'DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete * from TrailerMilesHours")
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    SQLz = "INSERT INTO TrailerMilesHours ( LogActivityDate, UnitID, UnitNumber, Miles, MilesLTD, Hours, HoursLTD, ObcDataStatus, OBCDataDesc ) " & _
           "SELECT CostcoTrailerData.LogActivityDate, Units.UnitID, Units.UnitUnitNum, CostcoTrailerData.MilesDelta, CostcoTrailerData.MILES, " & _
           "CostcoTrailerData.HoursDelta, CostcoTrailerData.Hours, CostcoTrailerData.ObcDataStatus, CostcoTrailerData.ObcDataStatusDesc " & _
           "FROM CostcoTrailerData INNER JOIN Units ON CostcoTrailerData.VIN = Units.UnitVIN;"
    'DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    'Call GetMainesTrailerData("C:\FAAS\\FleetImports\mainestrailerdata", 1)
End Sub
