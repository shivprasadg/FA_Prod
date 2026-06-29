Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim RST        As DAO.Recordset
Dim SQLz       As String
Dim i          As Integer
Dim Ccrl       As Control

Private Sub Form_Close()
    SQLz = "UPDATE Units "
    SQLz = SQLz & "SET ReportStartDate = UnitInServ "
    SQLz = SQLz & "WHERE ReportStartDate Is Null "
    SQLz = SQLz & "AND NOT UnitInServ IS NULL "
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
End Sub

Private Sub GrabVINStoDecode_Click()
    DoCmd.RunSQL ("Delete * From Unit_Vin_temp")
    SQLz = "INSERT INTO Unit_Vin_temp (ClientID, VIN, UnitNum, UnitYear) SELECT " & Me.ClientID & ",VIN, UnitNumber, UnitYear From FleetStudyImport"
    DoCmd.RunSQL (SQLz)
End Sub
Private Sub CheckVINsOnly_Click()
    'If Me.DecodeClient & "test" = "test" Then Exit Sub
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
    '  Me.DecodeClient.Requery
    Call SetEngineType(True)
End Sub




Private Sub SelectToDelete_Click()
    For i = 0 To UnitsFormsList.ListCount - 1
        UnitsFormsList.Selected(i) = True
    Next i
End Sub
Private Sub DeleteUnits_Click()
    'If Me.MLOrg = "FA" Then
    'MsgBox "Sorry, You may not delete LEASED units", vbCritical, "Danger"
    'Exit Sub
    'End If
    For Each ItemIndex In Me.UnitsFormsList.ItemsSelected
        If DCount("UnitRef", "ChildParts", "[UnitRef]= " & Me.UnitsFormsList.ItemData(ItemIndex) & "") > 0 Then
            MsgBox "Sorry, aleast some of these units have child parts attached" & vbNewLine & "Delete aborted", vbCritical, "Whoa! You're going wild!"
            Exit Sub
        End If
    Next
    Dim Response, FormNameX As String
    Response = MsgBox("Are You Sure You Want To Delete These Units?", vbQuestion + vbYesNo + vbCritical, "Say hello to my little friend...")
    If Response = vbYes Then
        For Each ItemIndex In Me.UnitsFormsList.ItemsSelected
            SQLz = "Delete * From Units Where UnitId = " & Me.UnitsFormsList.ItemData(ItemIndex) & ";"
            DoCmd.SetWarnings False
            DoCmd.RunSQL (SQLz)
            DoCmd.SetWarnings True
            i = i + 1
        Next
        Me.UnitsFormsList.Requery
        Me.UnitList.Requery
    Else
        Exit Sub
    End If
End Sub
Private Sub GetAll_Click()
    For i = 0 To UnitList.ListCount - 1
        UnitList.Selected(i) = True

    Next i
End Sub

Private Sub MoveUnits_Click()
    If IsNull(Me.LocationList.Column(0)) Then
        MsgBox "You Must Select a Location", , "Need Loc"
        Exit Sub
    End If
    Dim ItemIndex As Variant
    Dim TempUnitNum, VinX As String
    Dim i      As Integer
    Dim IDSx   As String
    i = 0

    DoCmd.SetWarnings False
    Call MessageUser("Open", "Inserting Units", "collecting data...")
    For Each ItemIndex In Me.UnitList.ItemsSelected
        i = i + 1
        TempUnitNum = "xUnit#" & i
        SQLz = "INSERT INTO Units ( "
        SQLz = SQLz & "UnitUnitNum"
        SQLz = SQLz & ", UnitVIN"
        SQLz = SQLz & ", UnitType"
        SQLz = SQLz & ", UnitSubType"
        SQLz = SQLz & ", UnitYr"
        SQLz = SQLz & ", UnitMake"
        SQLz = SQLz & ", UnitModel"
        SQLz = SQLz & ", UnitInServ"
        SQLz = SQLz & ", Axle"
        SQLz = SQLz & ", UnitRent"
        SQLz = SQLz & ", TUnitLessorOwner"
        SQLz = SQLz & ", UnitFinalCost"
        SQLz = SQLz & ", UnitOdometer"
        SQLz = SQLz & ", UnitAddDt"
        SQLz = SQLz & ", UnitStatus"
        SQLz = SQLz & ", ReportStartDate"
        SQLz = SQLz & ", ReportEndDate"
        SQLz = SQLz & ", VendorREF"
        SQLz = SQLz & ", UnitGarageLocationRef"
        SQLz = SQLz & ", UnitBuildDate"
        SQLz = SQLz & ", GroupREF"
        SQLz = SQLz & ", ClientGroupId"
        SQLz = SQLz & ", ClientId"
        SQLz = SQLz & ", ObcUnitNumber"
        If Nz(Me.DivisionOwnerID, 0) > 0 Then SQLz = SQLz & ",DivisionOwnerID "
        SQLz = SQLz & ") "
        SQLz = SQLz & " SELECT "
        SQLz = SQLz & "  Nz(UnitNumber,'" & TempUnitNum & "') as UnitNum"
        SQLz = SQLz & ", FSW.VIN"
        SQLz = SQLz & ", Nz(FSW.[Type],'TBD') AS Type"
        SQLz = SQLz & ", Nz(FSW.SubType,'TBD') AS SubType"
        SQLz = SQLz & ", Nz(FSW.UnitYear,1901) AS MYear"
        SQLz = SQLz & ", Nz(FSW.UnitMake,'TBD') AS Make"
        SQLz = SQLz & ", Nz(FSW.UnitModel,'TBD') AS Model"
        SQLz = SQLz & ", FSW.InServiceDate As InserviceDate"
        SQLz = SQLz & ", Nz(Axle,'TBD') as Ax"
        SQLz = SQLz & ", Nz([UnitRent],0) AS Rent"
        SQLz = SQLz & ", Ownership"
        SQLz = SQLz & ", Nz(UnitCost,0) AS Cost"
        SQLz = SQLz & ", Odometer"
        SQLz = SQLz & ", Date() AS DateUnitAdded"
        SQLz = SQLz & ", 'T' AS Unitstatus"
        SQLz = SQLz & ", CDate('1/1/' & ([unityear]-1)) AS Start"
        SQLz = SQLz & ", CDate('12/31/2050') AS Enddate"
        SQLz = SQLz & ", 39 AS Vendor"
        SQLz = SQLz & ", Nz(FSW.LocationID, " & Me.LocationList.Column(0) & ") AS GarageLoc"
        SQLz = SQLz & ", CDate('1/1/' & ([unityear]-1)) AS BuildDate"
        SQLz = SQLz & ", " & Me.GroupList.Column(0) & " AS GroupID"
        SQLz = SQLz & ", " & Me.ClientGroupID & " As ClientGrp"
        SQLz = SQLz & ", " & Me.ClientID & " as ClientID"
        SQLz = SQLz & ", Nz(ObcUnitNumber,nz(UnitNumber,'')) as ObcUnitNum "
        If Nz(Me.DivisionOwnerID, 0) > 0 Then SQLz = SQLz & "," & Me.DivisionOwnerID & " As DivisionOwnerID "
        SQLz = SQLz & "FROM FleetStudyImport_Work As FSW Where FSW.VIN = '" & Me.UnitList.ItemData(ItemIndex) & "' ;"
        Call MessageUser("Update", "Inserting Unit " & i & " of " & Me.UnitList.ItemsSelected.Count, "Working...")
        DoCmd.RunSQL (SQLz)
        Dim tUnitX As Long
        tUnitX = DLookup("UnitID", "Units", "UnitVIN='" & Me.UnitList.ItemData(ItemIndex) & "'")
        SQLz = "Insert Into TrackedUnitsRentProfile ("
        SQLz = SQLz & "tUnitID, tAssignee, tLeaseStart, tLeaseEnd, tTerm, tLeaseType)"
        SQLz = SQLz & " SELECT " & tUnitX & " as tUnitID"
        SQLz = SQLz & ", FSW.BankLessor"
        SQLz = SQLz & ", FSW.LeaseStartDate"
        SQLz = SQLz & ", FSW.LeaseEndDate"
        SQLz = SQLz & ", FSW.LeaseTerm"
        SQLz = SQLz & ", FSW.LeaseType "
        SQLz = SQLz & "FROM FleetStudyImport_Work As FSW Where FSW.VIN = '" & Me.UnitList.ItemData(ItemIndex) & "'"
        SQLz = SQLz & " AND (nz(BankLessor,'') <> '' "
        SQLz = SQLz & " OR nz(LeaseStartDate,'') <> '' "
        SQLz = SQLz & " OR nz(LeaseEndDate,'') <> '' "
        SQLz = SQLz & " OR nz(LeaseTerm,'') <> '' "
        SQLz = SQLz & " OR nz(LeaseType,'') <> '');"
        DoCmd.RunSQL (SQLz)
    Next
    ' SQLz = "UPDATE Units SET UnitInServ = Switch([UnitYr]>2014,'1/1/' & ([UnitYr]),[UnitYr]<2015,'3/1/' & ([UnitYr]-1)) " & _
    '        "WHERE UnitInServ Is Null AND Units.GroupREF=" & Me.GroupList.Column(0) & ";"
    ' DoCmd.RunSQL (SQLz)
    SQLz = "UPDATE Units SET DrivingMpgGoal = 6.5, MpgGoal = 6.5 WHERE DrivingMpgGoal IS NULL and GroupREF=" & Me.GroupList.Column(0) & ";"
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    On Error GoTo fixme1
    Me.UnitList.Requery
    Me.UnitsFormsList.Requery
    Call MessageUser("Close")
    Exit Sub
fixme1:
    DoCmd.SetWarnings True
    Call SyncUnitsToLocationDivisions
    Call MessageUser("Close")
End Sub

Private Sub SearchClient_Change()
    SQLz = "SELECT Client, clientID FROM ClientsList " & _
        "WHERE Client Like '*" & Me.SearchClient.Text & "*' Or ClientShNm Like '*" & Me.SearchClient.Text & "*' ORDER BY Client;"
    Me.ClientList.RowSource = SQLz
End Sub


Private Sub AddNewCompany_Click()
    DoCmd.GotoRecord , , acNewRec
    Me.LockUnLock.caption = "Locked"
    Call LockIt
End Sub



Private Sub LockIt()
    If Me.LockUnLock.caption = "Locked" Then
        For Each Ccrl In Me.Controls
            If Me(Ccrl.Name).Tag = "Required" Then
                Me(Ccrl.Name).Locked = False
            End If
        Next Ccrl
        Me.LockUnLock.caption = "UnLocked"
    Else
        For Each Ccrl In Me.Controls
            If Me(Ccrl.Name).Tag = "Required" Then
                Me(Ccrl.Name).Locked = True
            End If
        Next Ccrl
        Me.LockUnLock.caption = "Locked"
    End If
End Sub

Private Sub ClientShNm_DblClick(Cancel As Integer)
    Call LockIt
End Sub

Private Sub CloseMe_Click()
    DoCmd.OpenForm "FAMainMenu"
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub CompanyName_DblClick(Cancel As Integer)
    Call LockIt
End Sub
Private Sub SearchVIN_Change()
    'Me.SearchVin = ""
    Me.SearchYear = ""
    Me.SearchMake = ""
    Me.SearchType = ""
    Me.SearchLoc = ""

    SQLz = "SELECT VIN, UnitNumber, UnitYear, UnitMake, UnitModel, [SubType] & ' ' & [Type] AS ClassType, Location, Division  " & _
        "FROM FleetStudyImport " & _
        "WHERE VIN Like '*" & Me.SearchVin.Text & "*' ORDER BY [SubType] & ' ' & [Type];"
    Me.UnitList.RowSource = SQLz
End Sub
Private Sub SearchYear_Change()
    Me.SearchVin = ""
    'Me.SearchYear = ""
    Me.SearchMake = ""
    Me.SearchType = ""
    Me.SearchLoc = ""
    SQLz = "SELECT VIN, UnitNumber, UnitYear, UnitMake, UnitModel, [SubType] & ' ' & [Type] AS ClassType, Location, Division  " & _
        "FROM FleetStudyImport " & _
        "WHERE UnitYear Like '*" & Me.SearchYear.Text & "*' ORDER BY [SubType] & ' ' & [Type];"
    'SQLz = "SELECT FleetStudyImport.Vin, FleetStudyImport.Unitnumber, FleetStudyImport.UnitYear AS [Year], FleetStudyImport.UnitMake AS Make, FleetStudyImport.UnitModel AS Model, " & _
     "FleetStudyImport.SubType, FleetStudyImport.Location FROM FleetStudyImport LEFT JOIN Units ON FleetStudyImport.Vin = Units.UnitVIN " & _
     "WHERE Units.UnitVIN Is Null and UnitYear Like '*" & Me.SearchYear.Text & "*';"
    Me.UnitList.RowSource = SQLz
End Sub
Private Sub SearchMake_Change()
    Me.SearchVin = ""
    Me.SearchYear = ""
    'Me.SearchMake = ""
    Me.SearchType = ""
    Me.SearchLoc = ""
    SQLz = "SELECT VIN, UnitNumber, UnitYear, UnitMake, UnitModel, [SubType] & ' ' & [Type] AS ClassType, Location, Division  " & _
        "FROM FleetStudyImport " & _
        "WHERE UnitMake Like '*" & Me.SearchMake.Text & "*' ORDER BY [SubType] & ' ' & [Type];"

    'SQLz = "SELECT FleetStudyImport.VIN, FleetStudyImport.Unitnumber, FleetStudyImport.UnitYear AS [Year], FleetStudyImport.UnitMake AS Make, FleetStudyImport.UnitModel AS Model, " & _
     "FleetStudyImport.SubType, FleetStudyImport.Location FROM FleetStudyImport LEFT JOIN Units ON FleetStudyImport.VIN = Units.UnitVIN " & _
     "WHERE Units.UnitVIN Is Null and FleetStudyImport.UnitMake Like '*" & Me.SearchMake.Text & "*';"
    Me.UnitList.RowSource = SQLz
End Sub
Private Sub SearchType_Change()
    Me.SearchVin = ""
    Me.SearchYear = ""
    Me.SearchMake = ""
    'Me.SearchType = ""
    Me.SearchLoc = ""
    SQLz = "SELECT VIN, UnitNumber, UnitYear, UnitMake, UnitModel, [SubType] & ' ' & [Type] AS ClassType, Location, Division  " & _
        "FROM FleetStudyImport " & _
        "WHERE [SubType] & ' ' & [Type] Like '*" & Me.SearchType.Text & "*' ORDER BY [SubType] & ' ' & [Type];"
    ' SQLz = "SELECT FleetStudyImport.VIN, FleetStudyImport.Unitnumber, FleetStudyImport.UnitYear AS [Year], FleetStudyImport.UnitMake AS Make, FleetStudyImport.UnitModel AS Model, " & _
      "FleetStudyImport.SubType, FleetStudyImport.Location FROM FleetStudyImport LEFT JOIN Units ON FleetStudyImport.VIN = Units.UnitVIN " & _
      "WHERE Units.UnitVIN Is Null and FleetStudyImport.SubType Like '*" & Me.SearchType.Text & "*';"
    Me.UnitList.RowSource = SQLz
End Sub
Private Sub SearchLoc_Change()
    Me.SearchVin = ""
    Me.SearchYear = ""
    Me.SearchMake = ""
    Me.SearchType = ""
    'Me.SearchLoc = ""
    SQLz = "SELECT VIN, UnitNumber, UnitYear, UnitMake, UnitModel, [SubType] & ' ' & [Type] AS ClassType, Location, Division  " & _
        "FROM FleetStudyImport " & _
        "WHERE Location Like '*" & Me.SearchLoc.Text & "*' ORDER BY [SubType] & ' ' & [Type];"

    '  SQLz = "SELECT FleetStudyImport.VIN, FleetStudyImport.Unitnumber, FleetStudyImport.UnitYear AS [Year], FleetStudyImport.UnitMake AS Make, FleetStudyImport.UnitModel AS Model, " & _
       "FleetStudyImport.SubType, FleetStudyImport.Location FROM FleetStudyImport LEFT JOIN Units ON FleetStudyImport.VIN = Units.UnitVIN " & _
       "WHERE Units.UnitVIN Is Null and Location Like '*" & Me.SearchLoc.Text & "*';"
    Me.UnitList.RowSource = SQLz
End Sub

Private Sub SearchDivision_Change()
    Me.SearchVin = ""
    Me.SearchYear = ""
    Me.SearchMake = ""
    Me.SearchType = ""
    Me.SearchLoc = ""
    SQLz = "SELECT VIN, UnitNumber, UnitYear, UnitMake, UnitModel, [SubType] & ' ' & [Type] AS ClassType, Location, Division " & _
        "FROM FleetStudyImport " & _
        "WHERE Division  Like '*" & Me.SearchDivision.Text & "*' ORDER BY Division;"
End Sub

Private Sub SubmitCompany_Click()
    On Error GoTo FixMe
    DoCmd.RunCommand acCmdSaveRecord
    Me.ClientList = Me.ClientID
    Call ClientList_Click
FixMe:
End Sub
Private Sub ClientList_Click()
    Me.Form.FilterOn = False
    Me.LeaseList.RowSource = ""
    Me.ScheduleList.RowSource = ""
    Me.GroupList.RowSource = ""
    Me.Form.RecordSource = "Select * From Clients Where ClientId=" & Me.ClientList & ";"
    ' Set RST = Me.RecordsetClone
    ' RST.FindFirst "clientID = " & Me.ClientList.Column(1) & ""
    ' Me.Bookmark = RST.Bookmark
    Call MakeWorkTable
    Set RST = Nothing
    SQLz = "SELECT DivisionID, DivisionName, ClientID " & _
        "FROM ClientDivisions " & _
        "WHERE ClientID=" & Me.ClientList & ";"
    Me.DivisionOwnerID.RowSource = SQLz

    SQLz = "SELECT MLID, MLNo, MLOrig, Client, MLDt, MLType " & _
        "From MstrLease WHERE Client = " & Me.ClientList & " ORDER BY MstrLease.MLNo;"
    Me.LeaseList.RowSource = SQLz
    SQLz = "SELECT LocationID, LocationType, LocationName as LocName, LCity as City, LState as State, Locationsn as Loc, LCounty, LCountry, LPhone, ClientRef " & _
        "FROM Locations Where ClientGroupID=" & Me.ClientGroupID & " AND LocationType= 'Garage' ORDER By LocationName ;"
    Me.LocationList.RowSource = SQLz
    Me.UnitList.Requery
End Sub
Public Function MakeWorkTable()
    If Nz(DLookup("ClientID", "FleetStudyImport_Work"), 0) <> Me.ClientList Then
        MessageUser "Open", "Setting up Unit List", "Cleaning out old data"
        DoCmd.SetWarnings False
        DoCmd.RunSQL ("Delete From FleetStudyImport_Work")
        MessageUser "Update", "Creating List of units", "Getting Excel data..."
        SQLz = "INSERT INTO FleetStudyImport_Work ( UnitNumber, OBCUnitNumber, VIN, UnitYear, UnitMake, UnitModel, Type, SubType, InServiceDate, " & _
            "Location, LeaseEndDate, Axle, UnitRent, LeaseStartDate, Unitcost, LeaseTerm, Odometer, OdometerDate, Ownership,Division, State, Plate, ReeferHours,BankLessor, LeaseType ) " & _
            "SELECT FS.UnitNumber, FS.OBCUnitNumber, FS.VIN, FS.UnitYear, FS.UnitMake, FS.UnitModel, FS.Type, FS.SubType, FS.InServiceDate, " & _
            "FS.Location , FS.LeaseEndDate, FS.axle, FS.UNITRENT, FS.LeaseStartDate, FS.Unitcost, FS.LeaseTerm, FS.Odometer, FS.OdometerDate, " & _
            "FS.Ownership, FS.Division, FS.State, FS.Plate, FS.ReeferHours ,FS.BankLessor, FS.LeaseType " & _
            "FROM FleetStudyImportBase AS FS;"
        DoCmd.RunSQL (SQLz)
        DoCmd.RunSQL ("Update FleetStudyImport_Work Set ClientID = " & Me.ClientID & ", ClientGroupID=" & Me.ClientGroupID & "")
        DoCmd.RunSQL ("UPDATE FleetStudyImportLocID INNER JOIN FleetStudyImport_Work " & _
            "ON FleetStudyImportLocID.VIN = FleetStudyImport_Work.VIN " & _
            "SET FleetStudyImport_Work.LocationID = [FleetStudyImportLocID].[LocationID];")
        DoCmd.SetWarnings True
        MessageUser "Close"
    End If
End Function

Private Sub AddMasterLease_Click()
    Me.MLNo = Left(Me.ClientShNm, 4) & "-TK-" & DatePart("m", Date) & "-" & DatePart("YYYY", Date)
    Me.MLOrg = "Portfolio"
    Me.MLDt = Date
    Me.MLType = 4
End Sub
Private Sub PostImport_Click()
    If Me.ScheduleList & "test" = "test" Or Me.LeaseList & "test" = "test" Then
        MsgBox "You must select a Master Lease and a Schedule First", vbInformation, "You're silly"
        Exit Sub
    End If
    If Me.LeaseList.Column(2) = "FA" Then Call SetUnitsInGroup(Me.ClientID, "FA") Else Call SetUnitsInGroup(Me.ClientID, "NotFA")
End Sub

Public Function SortListBox(CtrlNameX As String, ListBoxX As String, FieldNamex As String)
    Dim Ccrl   As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Name Like "*Toggle*" And Me(Ccrl.Name).Name <> CtrlNameX Then
            If CtrlNameX Like "*1" And Me(Ccrl.Name).Name Like "*1" Then Me(Ccrl.Name) = 0
            If CtrlNameX Like "*2" And Me(Ccrl.Name).Name Like "*2" Then Me(Ccrl.Name) = 0
        End If
    Next Ccrl
    SQLz = SplitString(Me(ListBoxX).RowSource, " Order By", 0) & " Order By " & FieldNamex & ""
    If Me(CtrlNameX) = 0 Then SQLz = SQLz & " Desc"
    Me(ListBoxX).RowSource = SQLz
End Function

Private Sub SyncLocationDivisions_Click()
    Call SyncUnitsToLocationDivisions
    MsgBox "Locations / Divisions Synced to the Divisions Assigned to the Units' Locations" & vbNewLine & "(if any needed to change)", vbInformation
End Sub

Private Sub ToggleSortClassType1_Click()
    Call SortListBox("ToggleSortClassType1", "UnitList", "[SubType] & ' ' & [Type]")
End Sub

Private Sub ToggleSortDivision1_Click()
    Call SortListBox("ToggleSortDivision1", "UnitList", "Division")
End Sub

Private Sub ToggleSortLocation1_Click()
    Call SortListBox("ToggleSortLocation1", "UnitList", "Location")
End Sub

Private Sub ToggleSortMake1_Click()
    Call SortListBox("ToggleSortMake1", "UnitList", "UnitMake")
End Sub
Private Sub ToggleSortModel1_Click()
    Call SortListBox("ToggleSortModel1", "UnitList", "UnitModel")
End Sub

Private Sub ToggleSortVIN1_Click()
    Call SortListBox("ToggleSortVIN1", "UnitList", "VIN")
End Sub

Private Sub ToggleSortYear1_Click()
    Call SortListBox("ToggleSortYear1", "UnitList", "UnitYear")
End Sub


Private Sub UpdateModel_Click()

    SQLz = "UPDATE Units INNER JOIN "
    SQLz = SQLz & "(Units AS U INNER JOIN UnitVinDetails AS D ON U.UnitId = D.UnitID) "
    SQLz = SQLz & "ON Units.UnitID = U.UnitID "

    SQLz = SQLz & "SET Units.UnitModel = [D].[model] "

    SQLz = SQLz & "WHERE U.ClientID=" & Me.ClientID & ";"

    DoCmd.RunSQL (SQLz)
End Sub
Private Sub UnitNumToObcNum_Click()
    Call MessageUser("Open", "Setting OBC Numbers", "...This will be quick...", False)
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "UnitUnitNun_2_OBCUnitNumber", acViewNormal
    Call MessageUser("Update", "Setting Client and ClientGroupID", "...This will be even quicker...", True)
    DoCmd.OpenQuery "UnitsMissing_ClientGroupId", acViewNormal
    DoCmd.SetWarnings True
    Call MessageUser("Close")
End Sub



Private Sub UpdateMstLease_Click()
    SQLz = "UPDATE SchGrp SET MLDt = '" & Me.MLDt.Value & "', MLOrig = '" & Me.MLOrg.Value & "', MLNo = '" & Me.MLNo.Value & "' MLType ='" & Me.MLType.Value & "' " & _
        "WHERE MLID=" & Me.MLID & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Me.LeaseList.Requery
End Sub
Private Sub SubmitMasterLease_Click()
    If DCount("MLID", "MstrLease", "[MLNo]='" & Me.MLNo & "'") > 0 Then
        '  Response = MsgBox("Are You Sure You Want To Delete Schedule " & Me.ScheduleList.Column(1) & "", vbQuestion + vbYesNo + vbCritical, "Are You Crazy!?")
        ' If Response = vbYes Then
        MsgBox "Sorry, but there is already a Master Lease " & vbNewLine & "with the name of: '" & Me.MLNo & "'", vbCritical, "Cloning is illegal."
        SQLz = "UPDATE SchGrp SET MLDt = '" & Me.MLDt.Value & "', MLOrig = '" & Me.MLOrg.Value & "', MLNo = '" & Me.MLNo.Value & "' MLType ='" & Me.MLType.Value & " " & _
            "WHERE MLID=" & Me.MLID & ";"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
        Me.LeaseList.Requery
        Exit Sub
    Else
        SQLz = "INSERT INTO MstrLease ( Client, MLOrig, MLNo, MLDt, MLType ) " & _
            "VALUES ( " & Me.ClientID.Value & ", '" & Me.MLOrg.Value & "', '" & Me.MLNo.Value & "', #" & Date & "#, " & Me.MLType & " );"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
        Me.LeaseList.Requery
        Me.SchNo = "Track1"
        Me.FAID = "Trk-" & Me.ClientID & Left(Me.CompanyName, 3)
        Me.ShortDesc = Me.ClientShNm & " Tracked (" & DCount("UnitNumber", "FleetStudyImport") & ") Units"
        Me.MLID = DLookup("MLID", "MstrLease", "[MLNo]='" & Me.MLNo & "'")
        Me.ScheduleList.Requery
    End If
End Sub
Private Sub Leaselist_Click()
    Me.ScheduleList.RowSource = ""
    Me.GroupList.RowSource = ""
    Me.MLNo = Me.LeaseList.Column(1)
    Me.MLOrg = Me.LeaseList.Column(2)
    Me.MLDt = Me.LeaseList.Column(4)
    Me.MLType = Me.LeaseList.Column(5)
    SQLz = "Select SchID, SchNo as Schedule, SchFAID As FAID, SchMLRef,DateSchAdded, SchShortDesc "
    SQLz = SQLz & "FROM Schedule WHERE SchMLRef =" & Me.LeaseList.Column(0) & " "
    SQLz = SQLz & "Order BY CDbl(NumInString(Left(SchNo,3))) desc, SchNo desc;"
    Me.ScheduleList.RowSource = SQLz
End Sub
Private Sub DeleteMLease_Click()
    Dim UnitCnt As Integer
    If IsNull(Me.LeaseList.Column(2)) Then Exit Sub
    UnitCnt = DCount("AssetID", "vw_SixKeys", "[MLID]=" & Me.LeaseList.Column(0) & "")
    If UnitCnt > 0 Then    ' or Me.LeaseList.Column(2) = "FA" then
        MsgBox "You Cannot Delete Master Lease '" & Me.LeaseList.Column(1) & "'" & vbNewLine & "Because it contains " & UnitCnt & " units", vbCritical, "Danger Will Robertson!"
        Exit Sub
    End If
    Dim Response, FormNameX As String
    Response = MsgBox("Are You Sure You Want To Delete Lease " & Me.LeaseList.Column(1) & "", vbQuestion + vbYesNo + vbCritical, "Are You Crazy!?")
    If Response = vbYes Then

        SQLz = "Delete * From Schedule WHERE SchMLRef =" & Me.LeaseList.Column(0) & ";"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
        Me.SchNo = Null
        Me.FAID = Null
        Me.MLID = Null


        SQLz = "Delete  From MstrLease WHERE MLID =" & Me.LeaseList.Column(0) & ";"
        Me.LeaseList = 0
        Me.SchNo = Null
        Me.FAID = Null
        Me.MLID = Null
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True


    End If
    Me.LeaseList.Requery
    Me.ScheduleList.Requery
End Sub
Private Sub AddSchedule_Click()
    If IsNull(Me.LeaseList.Column(1)) Then
        MsgBox "You MUST Select A Master Lease", vbCritical, "No Short Cuts"
        Exit Sub
    End If
    Me.SchNo = "Track1"
    Me.FAID = "Trk-" & Me.ClientID & Left(Me.CompanyName, 3)
    Me.ShortDesc = Me.ClientShNm & " Tracked (" & DCount("UnitNumber", "FleetStudyImport") & ") Units"
    Me.MLID = Me.LeaseList.Column(0)
    Me.MLID = DLookup("MLID", "MstrLease", "[MLNo]='" & Me.MLNo & "'")
End Sub

Private Sub ScheduleList_Click()
    Me.SchNo = Me.ScheduleList.Column(1)
    FAID = Me.ScheduleList.Column(2)
    Me.ShortDesc = Me.ScheduleList.Column(5)
    Me.SchID = Me.ScheduleList.Column(0)
    Me.UnitsFormsList.RowSource = ""
    Me.MLID = DLookup("MLID", "MstrLease", "[MLNo]='" & Me.MLNo & "'")
    SQLz = "Select SGrpID, SGrpNO As GroupNo, LeaseType As GType, SGrpFAID, SgrpDesc, SGrpSchID FROM SchGrp WHERE SgrpSchID = " & ScheduleList & " ORDER BY SGrpNO;"
    Me.GroupList.RowSource = SQLz
    Me.GroupList = Null
    Me.groupNo = Null
    Me.groupType = Null
    Me.GroupDesc = Null
    Me.GrpSchID = Null
    Me.groupID = Null
End Sub
Private Sub UpdateSchedule_Click()
    If DCount("SchNo", "Schedule", "[SchMLRef] = " & Me.MLID & " And [SchNo]='" & Me.SchNo & "'") > 0 Then
        MsgBox "Sorry, but there is already a Schedule " & vbNewLine & "with the name of: '" & Me.SchNo & "'", vbCritical, "Cloning is illegal."
        Exit Sub
    Else
        SQLz = "UPDATE Schedule SET SchNO = '" & Me.SchNo & "', SchFAID = '" & Me.FAID & "', SchShortDesc = '" & Me.ShortDesc & "' " & _
            "WHERE SchID=" & Me.SchID & ";"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
    End If
    Me.ScheduleList.Requery
    Me.GroupList.Requery
End Sub
Private Sub SubmitSchedule_Click()

    If DCount("SchNo", "Schedule", "[SchMLRef] = " & Me.MLID & " And [SchNo]='" & Me.SchNo & "'") > 0 Then
        MsgBox "Sorry, but there is already a Schedule " & vbNewLine & "with the name of: '" & Me.SchNo & "'", vbCritical, "Cloning is illegal."
        Exit Sub
    Else
        SQLz = "INSERT INTO Schedule ( SchMLRef,SchNo, SchFAID, DateSchAdded, SchShortDesc ) " & _
            "VALUES ( " & Me.MLID.Value & ", '" & Me.SchNo.Value & "', '" & Me.FAID.Value & "', #" & Date & "#, '" & Me.ShortDesc & "' );"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
    End If
    Me.ScheduleList.Requery
    Me.GroupList.Requery
End Sub
Private Sub DeleteSchedule_Click()
    Dim UnitCnt As Integer
    If IsNull(Me.ScheduleList.Column(2)) Then Exit Sub
    UnitCnt = DCount("AssetID", "vw_SixKeys", "[SchID]=" & Me.ScheduleList.Column(0) & "")
    If UnitCnt > 0 Then    ' or Me.LeaseList.Column(2) = "FA" then
        MsgBox "You Cannot Delete Schedule '" & Me.ScheduleList.Column(1) & "'" & vbNewLine & "Because it contains " & UnitCnt & " units", vbCritical, "Danger Will Robertson!"
        Exit Sub
    End If
    Dim Response, FormNameX As String
    Response = MsgBox("Are You Sure You Want To Delete Schedule " & Me.ScheduleList.Column(1) & "", vbQuestion + vbYesNo + vbCritical, "Are You Crazy!?")
    If Response = vbYes Then

        SQLz = "Delete * From SchGrp WHERE SGrpSchID =" & Me.ScheduleList.Column(0) & ";"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
        Me.groupNo = Null
        Me.groupType = Null
        Me.GroupDesc = Null
        Me.GrpSchID = Null

        SQLz = "Delete * From Schedule WHERE SchID =" & Me.ScheduleList.Column(0) & ";"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
        Me.SchNo = Null
        Me.FAID = Null
        Me.MLID = Null
        Me.ShortDesc = Null
    End If
    Me.GroupList.Requery
    Me.ScheduleList.Requery
End Sub

Private Sub GroupList_Click()
    Me.groupID = Me.GroupList.Column(0)
    Me.groupNo = Me.GroupList.Column(1)
    Me.groupType = Me.GroupList.Column(2)
    Me.GroupDesc = Me.GroupList.Column(4)

    SQLz = "SELECT FleetStudyImport.VIN, FleetStudyImport.Unitnumber, FleetStudyImport.UnitYear AS [Year], FleetStudyImport.UnitMake AS Make, FleetStudyImport.UnitModel AS Model, " & _
        "FleetStudyImport.SubType, FleetStudyImport.Location FROM FleetStudyImport LEFT JOIN Units ON FleetStudyImport.VIN = Units.UnitVIN " & _
        "WHERE Units.UnitVIN Is Null ORDER BY FleetStudyImport.Location;"
    'Me.UnitList.RowSource = SQLz
    Me.UnitList.Requery
    SQLz = "SELECT Units.UnitID AS AssetId, Units.UnitVIN AS VIN, Units.UnitUnitNum AS UnitNo, Units.UnitMake, Units.UnitSubType, " & _
        "Locations.LocationName, Units.UnitYr FROM Units LEFT JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID " & _
        "WHERE GroupREF=" & Me.GroupList & " ORDER BY UnitID, UnitType, LocationName, UnitUnitNum;"
    Me.UnitsFormsList.RowSource = SQLz
    Exit Sub
FixMe:
End Sub
Private Sub AddGroup_Click()
    If IsNull(Me.ScheduleList.Column(2)) Then
        MsgBox "You MUST Select A Schedule", vbCritical, "No Short Cuts"
        Exit Sub
    End If
    Me.groupNo = "A"
    Me.groupType = "REPORTING"
    Me.GroupDesc = "Tracking Group"
    Me.GrpSchID = Me.ScheduleList.Column(0)
End Sub

Private Sub SubmitGroup_Click()
    If DCount("SGrpNO", "SchGrp", "[SGrpSchID] = " & Me.GrpSchID & " And [SGrpNO]='" & Me.groupNo & "'") > 0 Then
        MsgBox "Sorry, but there is already a Group " & vbNewLine & "with the name of: '" & Me.groupNo & "'", vbCritical, "I Said 'Cloning is illegal' :(."
        Exit Sub
    Else
        SQLz = "INSERT INTO SchGrp ( SGrpNO, LeaseType, SGrpFAID, GroupAddedDate, SgrpDesc, SGrpSchID  ) " & _
            "VALUES ( '" & Me.groupNo.Value & "', '" & Me.groupType.Value & "', '" & Me.FAID.Value & "', #" & Date & "#, '" & Me.GroupDesc & "', " & Me.GrpSchID & ");"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
    End If
    Me.GroupList.Requery
End Sub
Private Sub UpDateGroup_Click()
    If DCount("SGrpNO", "SchGrp", "[SGrpSchID] = " & Me.SchID & " And [SGrpNO]='" & Me.groupNo & "'") > 0 Then
        MsgBox "Sorry, but there is already a Group " & vbNewLine & "with the name of: '" & Me.groupNo & "'", vbCritical, "I Said 'Cloning is illegal' :(."
        Exit Sub
    Else
        SQLz = "UPDATE SchGrp SET SGrpNO = '" & Me.groupNo & "', LeaseType = '" & Me.groupType & "', SgrpDesc = '" & Me.GroupDesc & "' " & _
            "WHERE SGrpID=" & Me.groupID & ";"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
        Me.GroupList.Requery
    End If
End Sub
Private Sub DeleteGroup_Click()
    Dim UnitCnt As Integer
    If IsNull(Me.GroupList.Column(2)) Then Exit Sub
    UnitCnt = DCount("AssetID", "vw_SixKeys", "[SGrpId]=" & Me.GroupList.Column(0) & "")
    If UnitCnt > 0 Then    ' or Me.LeaseList.Column(2) = "FA" then
        MsgBox "You Cannot Delete Group '" & Me.GroupList.Column(1) & "'" & vbNewLine & "Because it contains " & UnitCnt & " units", vbCritical, "Don't Play With Matches!"
        Exit Sub
    End If
    Dim Response, FormNameX As String
    Response = MsgBox("Are You Sure You Want To Delete Group " & Me.GroupList.Column(1) & "", vbQuestion + vbYesNo + vbCritical, "Are You Crazy!?")
    If Response = vbYes Then

        SQLz = "Delete * From SchGrp WHERE SGrpID =" & Me.GroupList.Column(0) & ";"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
        Me.groupNo = Null
        Me.groupType = Null
        Me.GroupDesc = Null
        Me.GrpSchID = Null
    End If
    Me.GroupList.Requery

End Sub

Private Sub LocationList_Click()
    Me.LocationID = LocationList.Column(0)
    Me.LocName = LocationList.Column(2)
    Me.LocCity = LocationList.Column(3)
    Me.LocState = LocationList.Column(4)
    Me.LocZip = LocationList.Column(5)
    Me.LocCounty = LocationList.Column(6)
    Me.LocCountry = LocationList.Column(7)
    Me.LocPhone = LocationList.Column(8)

End Sub
Private Sub AddLocation_Click()
    Me.LocName = "Input Name"
    Me.LocCity = "Input City"
    Me.LocState = Null
    Me.LocZip = Null
    Me.LocCounty = Null
    Me.LocCountry = "USA"
    Me.LocPhone = Null
    Me.LocationID = Null
End Sub
Private Sub SubmitLocation_Click()

    SQLz = "INSERT INTO Locations ( LocationType, LocationName, LCity, LState, LZip, LCounty, LCountry, LPhone, ClientRef  ) " & _
        "VALUES ( 'Garage', '" & Me.LocName.Value & "', '" & Me.LocCity.Value & "', '" & Me.LocState & "', '" & Me.LocZip & "', '" & Me.LocCounty & "', '" & Me.LocCountry & "', '" & Me.LocPhone & "', " & Me.ClientID & " );"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    'End If
    Me.LocationList.Requery

End Sub
Private Sub DeleteLocation_Click()
    Dim UnitCnt As Integer
    If IsNull(Me.LocationList.Column(2)) Then Exit Sub
    UnitCnt = DCount("UnitId", "Units", "UnitGarageLocationRef=" & Me.LocationList.Column(0) & "")
    If UnitCnt > 0 Then
        MsgBox "You Cannot Delete Location '" & Me.LocationList.Column(2) & "'" & vbNewLine & "Because it contains " & UnitCnt & " units", vbCritical, "You'll shoot your eye out!"
        Exit Sub
    End If
    Dim Response, FormNameX As String
    Response = MsgBox("Are You Sure You Want To Delete Location " & Me.LocationList.Column(2) & "", vbQuestion + vbYesNo + vbCritical, "Are You for Real!?")
    If Response = vbYes Then

        SQLz = "Delete * From Locations WHERE LocationId =" & Me.LocationList.Column(0) & ";"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
        Me.LocationID = Null
        Me.LocationList.Requery
        Call SyncUnitsToLocationDivisions
    End If
End Sub
