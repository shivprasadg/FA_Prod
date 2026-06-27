Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz       As String
Dim RST        As DAO.Recordset
Public Sub ClearBoxes()
    Dim Ccrl   As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Medit" Then
            Me(Ccrl.Name).Value = ""
        End If
    Next Ccrl
End Sub

Private Sub ClearFilters_Click()
  '  Me.ReMarketingMassMod_subform1.Form.FilterOn = False
End Sub
Private Sub HideSubs()
    Me.ReMarketingMassMod_sub_SUFParked.Form.RecordSource = ""
    Me.ReMarketingMassMod_sub_BuyOutView.Form.RecordSource = ""
    Me.ReMarketingMassMod_sub_FMVPenalty.Form.RecordSource = ""
    Me.ReMarketingMassMod_sub_SoldInvoice.Form.RecordSource = ""
    Me.ReMarketingMassMod_sub_SUFParked.Visible = False
    Me.ReMarketingMassMod_sub_BuyOutView.Visible = False
    Me.ReMarketingMassMod_sub_FMVPenalty.Visible = False
    Me.ReMarketingMassMod_sub_SoldInvoice.Visible = False
    Me.ReMarketingMassMod_sub_Inspection2.Visible = False
    Me.RemarketingMassMod_sub_Titles.Visible = False

End Sub
Private Sub ShowSubs()
    Me.ReMarketingMassMod_sub_SUFParked.Visible = True
    Me!ReMarketingMassMod_sub_SUFParked.Form.ClearBoxes
    Me.ReMarketingMassMod_sub_SUFParked.Form.RecordSource = "RemarketingMassMod1"

    Me.ReMarketingMassMod_sub_BuyOutView.Visible = True
    Me!ReMarketingMassMod_sub_BuyOutView.Form.ClearBoxes
    Me.ReMarketingMassMod_sub_BuyOutView.Form.RecordSource = "RemarketingMassMod1"

    Me.ReMarketingMassMod_sub_FMVPenalty.Visible = True
    Me!ReMarketingMassMod_sub_FMVPenalty.Form.ClearBoxes
    Me.ReMarketingMassMod_sub_FMVPenalty.Form.RecordSource = "RemarketingMassMod1"

    Me.ReMarketingMassMod_sub_SoldInvoice.Visible = True
    Me.ReMarketingMassMod_sub_SoldInvoice.Form.RecordSource = "RemarketingMassMod1"
    Me.ReMarketingMassMod_sub_Inspection2.Visible = True
    Me.ReMarketingMassMod_sub_Inspection2.Form.RecordSource = "SELECT UnitInspections.* FROM UnitInspections " & _
                                                              "INNER JOIN UnitInsp_WorkTable " & _
                                                              "ON UnitInspections.InspnID = UnitInsp_WorkTable.InspnID " & _
                                                              "ORDER BY UnitInspections.AssetID, UnitInspections.InspNum DESC;"
    Me.RemarketingMassMod_sub_Titles.Visible = True
    Me.RemarketingMassMod_sub_Titles.Form.RecordSource = "RMKTUnitTitleLog_Work"
    Me.RemarketingMassMod_sub_Titles.Form.Requery
    'Me!ReMarketingMassMod_sub_Inspection2.Form.ClearBoxes
    ' Me.ReMarketingMassMod_sub_Inspection2.Form.RecordSource = "RemarketingMassMod1"
End Sub

Private Sub ClearClientSearch_Click()
    Me.SearchClient = ""
    Call SearchClient_Change
End Sub

Private Sub CloseMe_Click()
    DoCmd.OpenForm "FAMainMenu"
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub Form_Load()
    Me.SchNo = "Please select a client..."
    SQLz = "SELECT clientID, ClientShortName, CompanyName, Count(AssetID),ClientGroupID AS [Unit Count] " & _
           "From vw_SixKeys " & _
           "GROUP BY clientID, ClientShortName, CompanyName,ClientGroupID " & _
           "ORDER BY CompanyName;"
    Me.ClientList.RowSource = SQLz
    Me.SearchAsset = ""
    Me.SearchUnit = ""
    Me.SearchVin = ""
    Me.SearchStatus = ""
    Me.SearchMake = ""
    Me.SearchType = ""
    Call DelWorkTable
    Call ClearBoxes
    Me.LeaseList.RowSource = ""
    Me.ScheduleList.RowSource = ""
    Me.Units_List.RowSource = ""
    Me.UnitCount.Visible = False
    Call HideSubs
End Sub

Public Sub GetAll_Click()
    For i = 0 To Units_List.ListCount - 1
        Units_List.Selected(i) = True
    Next i
    Call GetSelectedUnits_Click
End Sub

Public Sub GetSelectedUnits_Click()
    Dim ItemIndex As Variant
    'On Error GoTo FixMe
    Call DelWorkTable
    DoCmd.SetWarnings False
    For Each ItemIndex In Me.Units_List.ItemsSelected
        SQLz = "INSERT INTO RMKT_MMod_Work (UnitID) SELECT " & Me.Units_List.ItemData(ItemIndex) & " AS UnitID;"
        DoCmd.RunSQL (SQLz)
     Next
        SQLz = "INSERT INTO UnitInsp_WorkTable (InspnID) SELECT Distinct InspnID From UnitInspections Where AssetID in (Select UnitID FRom RMKT_MMod_Work)"
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
    Call InsertMassRMKTUnits
    Call ShowSubs
    'Me.UnitCount = Me.BulkChange_SubForm.Form!UnitCount
    Call ClearBoxes
    Call GetUnitCounts
    Me.UnitCount.Visible = True
    ' Make Title Log Work table for new Mass-Mod Titles.
    SQLx = ""
    SQLx = "INSERT INTO RMKTUnitTitleLog_Work ( AssetID, VIN, TitleNumber, TitleSt, [Desc], LastDirection, UserID, LastEntry, LastEntryTypeId, UnitStatus) "
    SQLx = SQLx & "SELECT S.AssetID, S.VIN, H.TitleNumber, S.TitledState, S.UnitDescShortVer"
    SQLx = SQLx & ", IIf([EntryTypeId]=1,'IN',IIf([EntryTypeID]=2,'Out','-')) AS Dir, '" & FAUserName & "' AS [User]"
    SQLx = SQLx & ", nz(H.EntryReason & ' (' & H.UnitTitleLogDate &')','N/A')"
    SQLx = SQLx & ", EntryTypeId, [Status] "
    SQLx = SQLx & "FROM (vw_SixKeys AS S INNER JOIN Units AS U ON S.AssetID = U.UnitID) "
    SQLx = SQLx & "LEFT JOIN UnitTitleHistoryLastEntry AS H ON U.UnitID = H.AssetID "
    SQLx = SQLx & "WHERE (((S.AssetID) In (Select UnitID From RMKT_MMod_Work)));"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLx)
    DoCmd.SetWarnings True
FixMe:
    On Error GoTo 0
    Exit Sub
End Sub
Private Sub DelWorkTable()
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("DELETE * From RMKT_MMod_Work;")
    DoCmd.RunSQL ("DELETE * From UnitInsp_WorkTable;")
    DoCmd.RunSQL ("DELETE * From RMKTUnitTitleLog_Work;")
    DoCmd.SetWarnings True
End Sub
Private Sub ReQueryAll()
  '  Me.ReMarketingMassMod_subform1.Requery
  'Me!ReMarketingMassMod_Titles.Form.ClearBoxes
End Sub

Private Sub GetUnitCounts()
    Dim UnitX, UnitZ As String
    If ([Units_List].[ListCount] - 1) = 1 Then UnitX = "Unit " & ChrW(9658) Else UnitX = "Units " & ChrW(9658)
    If [Units_List].[ItemsSelected].[Count] = 1 Then UnitZ = "Unit " & ChrW(9658) Else UnitZ = "Units " & ChrW(9658)
    Me.GetAll.caption = ChrW(9660) & "All " & ([Units_List].ListCount - 1) & " Listed " & UnitX
    If ([Units_List].[ListCount] - 1) > 0 Then Me.GetAll.Enabled = True Else Me.GetAll.Enabled = False
    Me.GetSelectedUnits.caption = ChrW(9660) & "Selected (" & [Units_List].[ItemsSelected].[Count] & ") " & UnitZ
    If [Units_List].[ItemsSelected].[Count] > 0 Then Me.GetSelectedUnits.Enabled = True Else Me.GetSelectedUnits.Enabled = False
    If Me.PickModeToggle = True Then Call ReselectPickedUnits
End Sub
Public Sub ClientList_Click()
    SQLz = "SELECT MLID, MLNo, MLOrig, ClientID From vw_SixKeys " & _
           "WHERE [Status] <> 'B' " & _
           "GROUP BY MLID, MLNo, MLOrig, ClientID Having ClientID =" & Me.ClientList & " Order By MLOrig;"
    Me.LeaseList.RowSource = SQLz
    Me.ScheduleList.RowSource = ""
    Me.GroupList.RowSource = ""
    Me.Units_List.RowSource = ""
    Me.SchID = ""
    Me.SchNo = Me.ClientList.Column(1)
    Me.ClientGroupID = Me.ClientList.Column(4)
    Me.LocationList.RowSource = "SELECT LocationID,LocationSn From Locations WHERE LocationType in ('Garage','Repair') AND ClientGroupID=" & Me.ClientGroupID & " ORDER BY LocationSN;"
    Me.SearchLocation = ""

    'Me.ReMarketingMassMod_subform1.Form.RecordSource = ""
    Call HideSubs
    Call GetUnitCounts
End Sub
Public Sub Leaselist_Click()
    SQLz = "SELECT SchID, Schedule, MLID, Count(AssetID) AS Unt, ClientID, MLNo, MLOrig " & _
           "FROM vw_SixKeys WHERE [Status] <> 'B' " & _
           "GROUP BY SchID, Schedule, MLID, ClientID, MLNo, MLOrig HAVING MLID=" & Me.LeaseList & " " & _
           "ORDER BY CDbl(NumInString(Left(Schedule,3))), Schedule;"
    ' Sqlz = "SELECT * From ReMarketing_MLSchGrpUnits WHERE MLID=" & Me.LeaseList & ";"
    Me.ScheduleList.RowSource = SQLz
    Me.SchID = ""
    Me.SchNo = Me.ClientList.Column(1) & " MLNo: " & LeaseList.Column(1)
    Me.GroupList.RowSource = ""
    Me.Units_List.RowSource = ""
    Call HideSubs
    Call GetUnitCounts
End Sub

Private Sub LocationList_Click()
  SQLz = "SELECT SchID, sGrpID As GroupID, AssetID, UnitNum, VIN, [Status], Make, SubType, MYear, UnitGroup " & _
           "FROM vw_SixKeys WHERE [Status] <> 'B' " & _
           "AND MLOrig='FA' " & _
           "AND (LocationID=" & Me.LocationList & " OR ParkRepairID=" & Me.LocationList & ") " & _
           "ORDER BY AssetID;"

    Me.Units_List.RowSource = SQLz
    Me.SchNo = Me.ClientList.Column(1) & " | Location: " & Me.LocationList.Column(1)
    Me.SearchAsset = ""
    Me.SearchUnit = ""
    Me.SearchVin = ""
    Me.SearchStatus = ""
    Me.SearchMake = ""
    Me.SearchType = ""
    Call GetUnitCounts
End Sub

Private Sub OpenExitOptions_Click()
    DoCmd.close acForm, "EXITOptions"
    DoCmd.OpenForm "EXITOptions"
    If Me.GroupList.Column(0) & "test" <> "test" Then
        SchIDz = Me.SchID
        Forms!ExitOptions.ClientList = DLookup("ClientId", "vw_SixKeys", "SchID=" & SchIDz & "")
        Form_EXITOptions.ClientList_Click
        Forms!ExitOptions.LeaseList = DLookup("MLID", "vw_SixKeys", "SchID=" & SchIDz & "")
        Form_EXITOptions.Leaselist_Click
        Forms!ExitOptions.ScheduleList = SchIDz
        Form_EXITOptions.ScheduleList_Click
        Forms!ExitOptions.GroupList = Me.GroupList.Column(0)
        Form_EXITOptions.GroupList_Click
    End If
End Sub

Private Sub PickModeToggle_Click()
    Dim i As Integer
    For i = Me.Units_List.ListCount - 1 To 0 Step -1
        Me.Units_List.Selected(i) = 0
    Next i
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete From Pick_RMKTUnits_WorkTable")
    DoCmd.SetWarnings True
    If Me.PickModeToggle = True Then
        Me.PickModeToggle.caption = ChrW(9660) & "Pick On"
    Else
        Me.PickModeToggle.caption = "Pick Off"
    End If
    Call GetUnitCounts
End Sub

Private Sub RMKTTabControl_Change()
    Call SetSubData(Me.RMKTTabControl.Value)
End Sub
Public Function SetSubData(TabID As Integer)
    If TabID = 5 Then    ' Titles Tab
       Me.RemarketingMassMod_sub_Titles.Form.Requery
    End If
End Function

Public Sub ScheduleList_Click()
    Me.SchID = Me.ScheduleList
    Me.SchNo = Me.ClientList.Column(1) & " MLNo: " & LeaseList.Column(1) & " Sch: " & Me.ScheduleList.Column(2)

    SQLz = "SELECT SchID, sGrpID As GroupID, AssetID, UnitNum, VIN, [Status], Make, SubType, MYear, UnitGroup " & _
           "FROM vw_SixKeys WHERE [Status] <> 'B' " & _
           "AND SchID=" & Me.ScheduleList & " Order BY AssetID;"
    'Sqlz = "SELECT * " & _
     "FROM RemarketingMassModUnitList " & _
     "Where SchID=" & Me.SchID & " Order BY AssetID;"
    Me.Units_List.RowSource = SQLz

    SQLz = "SELECT SgrpID, UnitGroup, Count(AssetID) AS Unt, SchID " & _
           "FROM vw_SixKeys WHERE [Status] <> 'B' " & _
           "GROUP BY SgrpID, UnitGroup,SchID HAVING SchID=" & Me.ScheduleList & " Order BY UnitGroup;"
    '     Sqlz = "SELECT GroupId, UnitGroup, Count(AssetID) AS Units, SchID " & _
          "From RemarketingMassModUnitList " & _
          "GROUP BY GroupId, UnitGroup, SchID " & _
          "HAVING SchID=" & Me.ScheduleList & " Order BY UnitGroup; "
    Me.GroupList.RowSource = SQLz
    Call HideSubs
    '  SQLz = "Select * From RemarketingUnitsBySchedule Where SchID=" & Me.SchID & " Order BY UnitUnitNum;"
    '  Me.ReMarketingMassMod_subform1.Form.RecordSource = SQLz
    '  ReMarketingMassMod_subform1.Visible = True
    Me.SearchAsset = ""
    Me.SearchUnit = ""
    Me.SearchVin = ""
    Me.SearchStatus = ""
    Me.SearchMake = ""
    Me.SearchType = ""
    Call GetUnitCounts
End Sub
Public Sub GroupList_Click()
    SQLz = "SELECT SchID, sGrpID As GroupID, AssetID, UnitNum, VIN, [Status], Make, SubType, MYear, UnitGroup " & _
           "FROM vw_SixKeys WHERE [Status] <> 'B' " & _
           "AND sGrpID=" & Me.GroupList & " Order BY AssetID;"
    'Sqlz = "SELECT * " & _
     "FROM RemarketingMassModUnitList " & _
     "Where GroupID=" & Me.GroupList & " Order BY AssetID;"
    Me.Units_List.RowSource = SQLz
    Me.SchNo = Me.ClientList.Column(1) & " MLNo: " & LeaseList.Column(1) & " Sch: " & Me.ScheduleList.Column(2) & " Grp: " & GroupList.Column(1)
    Me.SearchAsset = ""
    Me.SearchUnit = ""
    Me.SearchVin = ""
    Me.SearchStatus = ""
    Me.SearchMake = ""
    Me.SearchType = ""
    Call GetUnitCounts
End Sub

Private Sub SearchAsset_Change()
'Me.SearchAsset = ""
    Me.SearchUnit = ""
    Me.SearchVin = ""
    Me.SearchStatus = ""
    Me.SearchMake = ""
    Me.SearchType = ""
    SQLz = "SELECT SchID, sGrpID As GroupID, AssetID, UnitNum, VIN, [Status], Make, SubType, MYear, UnitGroup " & _
           "FROM vw_SixKeys WHERE [Status] <> 'B' " & _
           "AND SchID=" & Me.SchID & " AND AssetId Like '*" & Me.SearchAsset.Text & "*' Order BY AssetID;"

    ' Sqlz = "SELECT * " & _
      "FROM RemarketingMassModUnitList " & _
      "Where SchID=" & Me.SchID & " AND AssetId Like '*" & Me.SearchAsset.Text & "*' Order BY AssetID;"
    Me.Units_List.RowSource = SQLz
    Call GetUnitCounts
End Sub

Private Sub SearchClient_Change()
    SQLz = "SELECT clientID, ClientShortName, CompanyName, Count(AssetID) AS [Unit Count],ClientGroupID " & _
           "From vw_SixKeys WHERE ClientShortName Like '*" & Me.SearchClient.Text & "*' Or CompanyName Like '*" & Me.SearchClient.Text & "*' " & _
           "GROUP BY clientID, ClientShortName, CompanyName,ClientGroupID " & _
           "ORDER BY CompanyName;"

    'Sqlz = "SELECT Client, clientID FROM ClientsList " & _
     "WHERE Client Like '*" & Me.SearchClient.Text & "*' Or ClientShNm Like '*" & Me.SearchClient.Text & "*' ORDER BY Client;"
    Me.ClientList.RowSource = SQLz
End Sub



Private Sub SearchLocation_Change()
    Me.SearchAsset = ""
    Me.SearchUnit = ""
    Me.SearchVin = ""
    Me.SearchStatus = ""
    Me.SearchMake = ""
    Me.SearchType = ""
    Me.LocationList.RowSource = "SELECT LocationID,LocationSn From Locations WHERE LocationType in ('Garage','Repair') " & _
                               " AND ClientGroupID=" & Me.ClientGroupID & " " & _
                               " AND LocationSn Like '*" & Me.SearchLocation.Text & "*' " & _
                               " AND LocationName Like '*" & Me.SearchLocation.Text & "*' ORDER BY LocationSN ;"
End Sub

Private Sub SearchUnit_Change()
    Me.SearchAsset = ""
    'Me.SearchUnit = ""
    Me.SearchVin = ""
    Me.SearchStatus = ""
    Me.SearchMake = ""
    Me.SearchType = ""
    SQLz = "SELECT SchID, sGrpID As GroupID, AssetID, UnitNum, VIN, [Status], Make, SubType, MYear, UnitGroup " & _
           "FROM vw_SixKeys WHERE [Status] <> 'B' " & _
           "AND SchID=" & Me.SchID & " AND UnitNum Like '*" & Me.SearchUnit.Text & "*' Order BY UnitNum;"

    Me.Units_List.RowSource = SQLz
    Call GetUnitCounts
End Sub
Private Sub SearchVIN_Change()
    Me.SearchAsset = ""
    Me.SearchUnit = ""
    'Me.SearchVin = ""
    Me.SearchStatus = ""
    Me.SearchMake = ""
    Me.SearchType = ""
    SQLz = "SELECT SchID, sGrpID As GroupID, AssetID, UnitNum, VIN, [Status], Make, SubType, MYear, UnitGroup " & _
           "FROM vw_SixKeys WHERE [Status] <> 'B' " & _
           "AND SchID=" & Me.SchID & " AND VIN Like '*" & Me.SearchVin.Text & "*' Order BY VIN;"

    Me.Units_List.RowSource = SQLz
    Call GetUnitCounts
End Sub
Private Sub SearchStatus_Change()
    Me.SearchAsset = ""
    Me.SearchUnit = ""
    Me.SearchVin = ""
    'Me.SearchStatus = ""
    Me.SearchMake = ""
    Me.SearchType = ""
    SQLz = "SELECT SchID, sGrpID As GroupID, AssetID, UnitNum, VIN, [Status], Make, SubType, MYear, UnitGroup " & _
           "FROM vw_SixKeys WHERE [Status] <> 'B' " & _
           "AND SchID=" & Me.SchID & " AND [Status] Like '*" & Me.SearchStatus.Text & "*' Order BY AssetID;"
    '  Sqlz = "SELECT * " & _
       "FROM RemarketingMassModUnitList " & _
       "Where SchID=" & Me.SchID & " and S Like '*" & Me.SearchStatus.Text & "*' Order BY AssetID;"
    Me.Units_List.RowSource = SQLz
    Call GetUnitCounts
End Sub
Private Sub SearchMake_Change()
    Me.SearchAsset = ""
    Me.SearchUnit = ""
    Me.SearchVin = ""
    Me.SearchStatus = ""
    ' Me.SearchMake = ""
    Me.SearchType = ""
    SQLz = "SELECT SchID, sGrpID As GroupID, AssetID, UnitNum, VIN, [Status], Make, SubType, MYear, UnitGroup " & _
           "FROM vw_SixKeys WHERE [Status] <> 'B' " & _
           "AND SchID=" & Me.SchID & " AND Make Like '*" & Me.SearchMake.Text & "*' Order BY AssetID;"
    'Sqlz = "SELECT * " & _
     "FROM RemarketingMassModUnitList " & _
     "Where SchID=" & Me.SchID & " AND Make Like '*" & Me.SearchMake.Text & "*' Order BY AssetID;"
    Me.Units_List.RowSource = SQLz
    Call GetUnitCounts
End Sub
Private Sub SearchType_Change()
    Me.SearchAsset = ""
    Me.SearchUnit = ""
    Me.SearchVin = ""
    Me.SearchStatus = ""
    Me.SearchMake = ""
    'Me.SearchType = ""
    SQLz = "SELECT SchID, sGrpID As GroupID, AssetID, UnitNum, VIN, [Status], Make, SubType, MYear, UnitGroup " & _
           "FROM vw_SixKeys WHERE [Status] <> 'B' " & _
           "AND SchID=" & Me.SchID & " AND SubType Like '*" & Me.SearchType.Text & "*' Order BY AssetID;"
    ' Sqlz = "SELECT * " & _
      "FROM RemarketingMassModUnitList " & _
      "Where SchID=" & Me.SchID & " AND CabType Like '*" & Me.SearchType.Text & "*' Order BY AssetID;"
    Me.Units_List.RowSource = SQLz
    Call GetUnitCounts
End Sub
Private Sub ClearLabel_Click()
    Me.SearchAsset = ""
    Me.SearchUnit = ""
    Me.SearchVin = ""
    Me.SearchStatus = ""
    Me.SearchMake = ""
    Me.SearchType = ""
    Me.SearchVin.SetFocus
    Call SearchVIN_Change
End Sub

Public Sub Units_List_Click()
    If Me.PickModeToggle = True Then
        DoCmd.SetWarnings False
        If Coalesce(Me.Units_List.Column(0), 0) > 0 Then DoCmd.RunSQL ("Insert into Pick_RMKTUnits_WorkTable(UnitID) SELECT " & Me.Units_List.Column(2) & ";")
        DoCmd.SetWarnings True
        'Call ReselectPickedUnits
    End If
    Call GetUnitCounts
End Sub
Public Sub ReselectPickedUnits()
    Dim i, x As Integer
    Dim pUnits As String
    Dim RST As DAO.Recordset
    SQLz = "Select * From Pick_RMKTUnits_WorkTable Order by UnitID Desc;"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    On Error Resume Next
    RST.MoveLast
    RST.MoveFirst
    On Error GoTo 0
    If RST.RecordCount > 0 Then
        For x = 1 To RST.RecordCount
            For i = Me.Units_List.ListCount - 1 To 0 Step -1
                If Me.Units_List.ItemData(i) = CStr(RST!UnitID) Then Me.Units_List.Selected(i) = 1
            Next i
            RST.MoveNext
        Next x
    End If
    RST.close
    Set RST = Nothing
End Sub


Private Sub Units_List_DblClick(Cancel As Integer)
    On Error GoTo FixMe
    Set RST = Me.ReMarketingMassMod_sub_SUFParked.Form.RecordsetClone
    RST.FindFirst "UnitVIN = '" & Me.Units_List.Column(5) & "'"
    Me.ReMarketingMassMod_sub_SUFParked.Form.Bookmark = RST.Bookmark
    Set RST = Me.ReMarketingMassMod_sub_BuyOutView.Form.RecordsetClone
    RST.FindFirst "UnitVIN = '" & Me.Units_List.Column(5) & "'"
    Me.ReMarketingMassMod_sub_BuyOutView.Form.Bookmark = RST.Bookmark
    Set RST = Me.ReMarketingMassMod_sub_FMVPenalty.Form.RecordsetClone
    RST.FindFirst "UnitVIN = '" & Me.Units_List.Column(5) & "'"
    Me.ReMarketingMassMod_sub_FMVPenalty.Form.Bookmark = RST.Bookmark
    Set RST = Me.ReMarketingMassMod_sub_SoldInvoice.Form.RecordsetClone
    RST.FindFirst "UnitVIN = '" & Me.Units_List.Column(5) & "'"
    Me.ReMarketingMassMod_sub_SoldInvoice.Form.Bookmark = RST.Bookmark
    Set RST = Me.ReMarketingMassMod_sub_Inspection2.Form.RecordsetClone
    RST.FindFirst "UnitVIN = '" & Me.Units_List.Column(5) & "'"
    Me.ReMarketingMassMod_sub_Inspection2.Form.Bookmark = RST.Bookmark
    Set rs = Nothing
FixMe:
End Sub
