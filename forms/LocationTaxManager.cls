Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz As String
Dim SQLx As String

Private Sub ClearSeachBoxText_Click()
    Call ClearSearches
End Sub
Private Sub ClearSeachBoxText2_Click()
    Call ClearSearches
End Sub

Private Sub CloseMe_Click()
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub ExcelReport_Click()
    Dim XLS, Header, SQLy As String
    Header = "Location Taxes Report: " & Date
    SQLy = SplitString(Me.LocationsList.RowSource, ", LT.StateSales", 0) & ",0 as [Other Tax], LT.StateSales" & SplitString(Me.LocationsList.RowSource, ", LT.StateSales", 1)
    XLS = "|A,DD3|A,Wrap|A,W18|B,TX|B,W30|E,BC|G,TC|H,BC|I,SR|J,P3|K,P3|L,cLabel:3:Tax Option|L,W08|L,TC|S,TC|M,cLabel:3:Loc Taxes|M,P3"
    XLS = XLS & "|N,FxAll=K4-M4|N,cColor:3:Orange|N,W10|O,cLabel:3:State Tax|O,P3|P,cLabel:3:County Tax|P,cLabel:3:County Tax"
    XLS = XLS & "|P,P3|Q,P3|Q,cLabel:3:City Tax|R,cLabel:3:State Rent Receipt|R,BC|R,W08|"
    SendReport2Excel SQLy, "Loc Taxes", Header, True, XLS, 0, True
End Sub

Private Sub Form_Current()
    Dim ClientX
    'ClientX = Nz(Me.clientID, 0)
    Me.SearchAssetId = ""
    Me.SearchUnitNum = ""
    Me.SearchVin = ""
    Me.SearchStatus = ""
    SQLz = "SELECT DivisionID, DivisionName, ClientID From ClientDivisions " & _
           "WHERE Deleted=False and ClientID=" & ClientX & " ORDER BY DivisionName"
    Me.DivisionID.RowSource = SQLz
End Sub

Public Sub Form_Load()
    SQLx = "SELECT LT.Client"
    SQLx = SQLx & ", LT.LocationName"
    SQLx = SQLx & ", LT.LocationSN"
    SQLx = SQLx & ", LT.LCity"
    SQLx = SQLx & ", LT.LState"
    SQLx = SQLx & ", LT.LCounty"
    SQLx = SQLx & ", LT.LZip"
    SQLx = SQLx & ", Format(LT.LUnits,'#,##0') AS LUnits"
    SQLx = SQLx & ", Format(LT.FAUnits,'#,##0') AS FAUnits"
    SQLx = SQLx & ", Format(LT.Distro,'Percent') AS Distro"
    SQLx = SQLx & ", Format(LT.LTaxRate,'Percent') AS LTaxRate"
    SQLx = SQLx & ", LT.LTaxOption"
    SQLx = SQLx & ", Format(LT.LocTaxes,'Percent') AS LocTaxes"
    SQLx = SQLx & ", LT.StateSales"
    SQLx = SQLx & ", LT.CountySales"
    SQLx = SQLx & ", LT.CitySales"
    SQLx = SQLx & ", LT.StateRReceipt"
    SQLx = SQLx & ", LT.LStateID"
    SQLx = SQLx & ", LT.LCountyID"
    SQLx = SQLx & ", LT.LCityID"
    SQLx = SQLx & ", LT.LocationID"
    SQLx = SQLx & ", LT.ClientGroupId"
    SQLx = SQLx & ", LT.clientID"
    SQLx = SQLx & " FROM vw_LocationDistribution_Taxes as LT "

  '  SQLx = "SELECT Client, LocationName, LCity, LState, LCounty, LZip, LUnits, FAUnits, Distro, LTaxRate, LTaxOption, LocTaxes, " & _
        "StateSales, CountySales, CitySales, StateRReceipt, LStateID, LCountyID, LCityID, LocationID, ClientGroupId, clientID " & _
        "FROM LocationDistroTaxes "
    Call ClearToggles(0)
    Call ClearSearches
End Sub
Private Sub ClearSearches(Optional SearchBox As String)

    Dim Ccrl  As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "SearchX" Then
            If Not IsNull(SearchBox) Then
                If SearchBox <> Me(Ccrl.Name).Name Then Me(Ccrl.Name) = ""
            Else
                Me(Ccrl.Name) = ""
            End If
        End If
    Next Ccrl
    Me.LocationsList = ""
End Sub

Private Sub LCity_Click()
    Dim CityReply As String
    Dim NumRecords As Integer
    If Me.LCity Like "*Unkn*" Then
        Me.LCity = "Unknown"
        Me.LCityID = 0
        GoTo SaveMe
    End If
    CityReply = LookUpByCity(Me.LCity, Me.LState)
    If CityReply = "" Then NumRecords = 0 Else NumRecords = CInt(SplitString(CityReply, "|", 0))
    If Me.LCounty <> CheckCounty(Nz(Me.LCounty, 0), Me.LState, Me.LCity) Then
        If Me.LCounty <> "Unknown" Then
            MsgBox "The city '" & Me.LCity & "'" & vbNewLine & "is not normaly listed to be within " & Me.LCounty & " county" & vbNewLine & "(the County was corrected accordingly)", vbInformation, AtlaasVersion
            Me.LCounty.SetFocus
        End If
    End If
    If NumRecords = 0 Then
        MsgBox "You've entered a City that is not in the ATLAAS Core, Would you like to add it?", vbCritical, AtlaasVersion
        Exit Sub
    End If
    Me.LCityID = Me.LCity.Column(1)
    Me.LCounty = Me.LCity.Column(2)
    Me.LCountyID = Me.LCity.Column(3)
    Me.LCity.RowSource = SplitString(CityReply, "|", 4)
    Me.LCounty.RowSource = SplitString(CityReply, "|", 5)
SaveMe:
    If Me.Form.Dirty Then Me.Form.Dirty = False
    Call LCounty_Click
End Sub

Private Sub LCounty_Click()
    Dim CountyReply As String
    Dim NumRecords As Integer
    If Me.LCounty Like "*Unkn*" Then
        Me.LCounty = "Unknown"
        Me.LCountyID = 0
        GoTo SaveMe
    End If
    CountyReply = LookUpByCounty(Nz(Me.LCounty, ""), Me.LState)
    If CountyReply = "" Then NumRecords = 0 Else NumRecords = CInt(SplitString(CountyReply, "|", 0))
    If NumRecords = 0 Then
        MsgBox "You've entered a County that is not in the ATLAAS Core, Would you like to add it?", vbCritical, AtlaasVersion
        Exit Sub
    End If
    Me.LCity.RowSource = SplitString(CountyReply, "|", 4)
    Me.LCounty.RowSource = SplitString(CountyReply, "|", 5)
    If NumRecords > 1 And Nz(Me.LCity, "Unknown") = "Unknown" Then
        MsgBox "There is more than 1 City for this County, Please make sure to select the Proper City", , AtlaasVersion
    End If
    If Nz(Me.LCity, "Unknown") = "Unknown" Or Me.LCity = "tbd" Then Me.LCity = SplitString(CountyReply, "|", 1)
    Me.LCountyID = Me.LCounty.Column(1)
    If CheckCity(Me.LCounty, Me.LState, Me.LCity) <> Me.LCity Then
        If Me.LCity <> "Unknown" Then
            MsgBox "The city '" & Me.LCity & "'" & vbNewLine & "is not normaly listed to be within " & Me.LCounty & " county" & vbNewLine & "(the city was modified accordingly)", vbCritical, AtlaasVersion
            Me.LCity.SetFocus
        End If
    End If
SaveMe:
    If Me.Form.Dirty Then Me.Form.Dirty = False
    If Me.LTaxRate & 1 = 1 Then Me.LTaxRate = Nz(DLookup("BaseTaxRate", "TaxStatesBase", "isnull(EndsOn) And TaxTypeID=1 And TaxStatesID=" & Me.LStateID), 0) + _
        Nz(DLookup("BaseTaxRate", "TaxCountiesBase", "isnull(EndsOn) And TaxTypeID=1 And TaxCountyID=" & Me.LCountyID), 0) + _
        Nz(DLookup("BaseTaxRate", "TaxCitiesBase", "isnull(EndsOn) And TaxTypeID=1 And TaxCityID=" & Me.LCityID), 0)
    If Me.Form.Dirty Then Me.Form.Dirty = False
    Me.LocationsList.Requery
End Sub

Private Sub LookupStateTaxes_Click()
    If Me.LState & "test" = "test" Or Me.LZip & "test" = "test" Then
        MsgBox "An Address must have both a State and a Zip," & vbNewLine & "in order to perform a 'Look Up Tax (Zip)'", vbCritical, AtlaasVersion
        Exit Sub
    End If
    Call LookupSaleTaxStates(LZip, LTaxRate, LStateID, LCountyID, LCityID, LocationID, Me.LCounty)
    Me.Form.Requery
    Me.LocationsList.Requery
End Sub

Private Sub LookUpTaxZip_Click()
    If Me.LState & "test" = "test" Or Me.LZip & "test" = "test" Then
        MsgBox "An Address must have both a State and a Zip," & vbNewLine & "in order to perform a 'Look Up Tax (Zip)'", vbCritical, AtlaasVersion
        Exit Sub
    End If
    Call LookUpTaxByZip2Tax(LZip, LTaxRate, LStateID, LCountyID, LCityID, LocationID)
    Me.Form.Requery
    Me.LocationsList.Requery
End Sub

Private Sub LState_Click()
    If Nz(Me.LState, "XX") = "XX" Then
        Me.LCity.RowSource = ""
        Me.LCounty.RowSource = ""
        Exit Sub
    End If

    If Me.LState.Column(1) Like "*RR*" Then Me.LTaxOption = "RR"
    'Me.LocationTax.Requery
    Dim StateReply As String
    Dim NumRecords As Integer
    Dim Msgx, ResposeX As String
    StateReply = Nz(LookUpByState(Me.LState), "0|")
    NumRecords = CInt(SplitString(StateReply, "|", 0))
    Me.LStateID = Me.LState.Column(3)
    Me.LCounty.RowSource = "SELECT County, CountyID From vw_CCSZ_Lookup WHERE StateID =" & Me.LStateID & " GROUP BY County, CountyID ORDER BY County;"
    Me.LCity.RowSource = "SELECT City, CityID, County, CountyID From vw_CCSZ_Lookup WHERE StateID =" & Me.LStateID & " GROUP BY City, CityID, County, CountyID ORDER BY City, County;"
    If DCount("CityID", "vw_CCSZ_Lookup", "StateID=" & Me.LStateID & " AND City='" & Me.LCity & "'") = 0 Then Msgx = Msgx & "<br />The City '<b>" & Me.LCity & "</b>' is not within the State of <b> " & Me.LState & "</b>"
    If DCount(Nz(Me.LCountyID, 0), "vw_CCSZ_Lookup", "StateID=" & Me.LStateID & " AND CountyID =" & Nz(Me.LCountyID, 0)) = 0 Then Msgx = Msgx & "<br />The County '<b>" & Me.LCounty & "</b>' is not within the Sate of <b> " & Me.LState & "</b>"

    If Msgx <> "" And Nz(Me.LCity, "") <> "" Then
        Msgx = "There is a address conflict:" & Msgx
        Msgx = Msgx & "<br /><br />If you press 'Proceed' ATLAAS will mark the unmatched fields to <b>'Unknown'</b><br />You than can select the proper values, afterwards"
        ResposeX = Dialog.RichBox(Msgx, vbCustom + vbExclamation, AtlaasVersion, , , 0, 1, 0, 0, "Undo", "Leave it", "Proceed")
        If ResposeX = 101 Then Me.Undo
        ' If ResposeX = vbBt2 Then
        If ResposeX = 103 Then
            If Msgx Like "*The City*" Then Me.LCity = "Unknown"
            If Msgx Like "*The City*" Then Me.LCityID = 0
            If Msgx Like "*County*" Then Me.LCounty = "Unknown"
            If Msgx Like "*County*" Then Me.LCountyID = 0
        End If
        If Me.Form.Dirty Then Me.Form.Dirty = False
    End If

    If NumRecords = 0 Then
        Me.LCity.RowSource = ""
        Me.LCounty.RowSource = ""
        Exit Sub
    End If
    If Me.Form.Dirty Then Me.Form.Dirty = False
End Sub

Private Sub OpenClient_Click()
    If Me.clientID = "" Or IsNull(Me.clientID) Then Exit Sub
    DoCmd.OpenForm "ClientsView", , , , , , Me.clientID
    DoCmd.SelectObject acForm, "ClientsView"
End Sub

Private Sub ResetForm_Click()
    Call ClearToggles(0)
    Call ClearSearches
    SQLx = "SELECT Client"
    SQLx = SQLx & ", LocationName"
    SQLx = SQLx & ", LocationSN"
    SQLx = SQLx & ", LCity"
    SQLx = SQLx & ", LState"
    SQLx = SQLx & ", LCounty"
    SQLx = SQLx & ", LZip"
    SQLx = SQLx & ", LUnits"
    SQLx = SQLx & ", FAUnits"
    SQLx = SQLx & ", Distro"
    SQLx = SQLx & ", LTaxRate"
    SQLx = SQLx & ", LTaxOption"
    SQLx = SQLx & ", LocTaxes"
    SQLx = SQLx & ", StateSales"
    SQLx = SQLx & ", CountySales"
    SQLx = SQLx & ", CitySales"
    SQLx = SQLx & ", StateRReceipt"
    SQLx = SQLx & ", LStateID"
    SQLx = SQLx & ", LCountyID"
    SQLx = SQLx & ", LCityID"
    SQLx = SQLx & ", LocationID"
    SQLx = SQLx & ", ClientGroupId"
    SQLx = SQLx & ", ClientID "
    SQLx = SQLx & " FROM LocationDistroTaxes "
    SQLz = SQLx & " Order By Client, LUnits Desc"
    Me.LocationsList.RowSource = SQLz

End Sub

Private Sub SearchALL_Button_Click()
    SQLz = "SELECT Client"
    SQLz = SQLz & ", LocationName"
    SQLz = SQLz & ", LocationSN"
    SQLz = SQLz & ", LCity"
    SQLz = SQLz & ", LState"
    SQLz = SQLz & ", LCounty"
    SQLz = SQLz & ", LZip"
    SQLz = SQLz & ", LUnits"
    SQLz = SQLz & ", FAUnits"
    SQLz = SQLz & ", Distro"
    SQLz = SQLz & ", LTaxRate"
    SQLz = SQLz & ", LTaxOption"
    SQLz = SQLz & ", LocTaxes"
    SQLz = SQLz & ", StateSales"
    SQLz = SQLz & ", CountySales"
    SQLz = SQLz & ", CitySales"
    SQLz = SQLz & ", StateRReceipt"
    SQLz = SQLz & ", LStateID"
    SQLz = SQLz & ", LCountyID"
    SQLz = SQLz & ", LCityID"
    SQLz = SQLz & ", LocationID"
    SQLz = SQLz & ", ClientGroupId"
    SQLz = SQLz & ", clientID "
    SQLz = SQLz & " FROM LocationDistroTaxes WHERE LUnits > 0 "
    If Me.SearchClient <> "" Then SQLz = SQLz & " AND Client Like'*" & Me.SearchClient & "*' "
    If Me.SearchLocation <> "" Then SQLz = SQLz & "AND LocationName Like '*" & Me.SearchLocation & "*' "
    If Me.SearchShort <> "" Then SQLz = SQLz & "AND LocationName Like '*" & Me.SearchShort & "*' "
    If Me.SearchLCity <> "" Then SQLz = SQLz & " AND LCity Like '*" & Me.SearchLCity & "*' "
    If Me.SearchLState <> "" Then SQLz = SQLz & " AND LState Like '*" & Me.SearchLState & "*' "
    If Me.SearchLCounty <> "" Then SQLz = SQLz & " AND LCounty Like'*" & Me.SearchLCounty & "*' "
    If Me.SearchLZip <> "" Then SQLz = SQLz & " AND LZip Like '*" & Me.SearchLZip & "*' "
    SQLz = SQLz & " ORDER BY Client, LUnits Desc, LocationName;"
    Me.LocationsList.RowSource = SQLz
    ' End If
End Sub

Private Sub ClientList_Click()
    Me.SearchClient = Me.ClientList.Column(1)
    Call SearchClient_Button_Click
End Sub

Private Sub SearchAssetId_Change()
    Call SearchUnitsList("AssetID")
End Sub

Private Sub SearchStatus_Change()
    Call SearchUnitsList("Status")
End Sub

Private Sub SearchUnitNum_Change()
    Call SearchUnitsList("UnitNum")
End Sub

Private Sub SearchVIN_Change()
    Call SearchUnitsList("VIN")
End Sub
Public Function SearchUnitsList(sField As String)
    If sField <> "AssetID" Then Me.SearchAssetId = ""
    If sField <> "VIN" Then Me.SearchVin = ""
    If sField <> "UnitNum" Then Me.SearchUnitNum = ""
    If sField <> "Status" Then Me.SearchStatus = ""
    If Len(Me("Search" & sField).Text) >= 1 And Len(Me("Search" & sField).Text) < 3 Then Exit Function
  '  If Len(Me("Search" & sField).Text) = 0 Then
  '      Me.UnitsList.RowSource = ""
  '      Exit Function
  '  End If
    SQLz = SplitString(Me.UnitsList.RowSource, " AND MLOrig = 'FA'", 0)
    SQLz = SQLz & " AND MLOrig = 'FA' AND " & sField & " Like '*" & Me("Search" & sField).Text & "*' Order BY " & sField & ";"
    Me.UnitsList.RowSource = SQLz
End Function
Private Sub SearchClient_Button_Click()
    Call ClearSearches("SearchClient")
    SQLz = SplitString(SQLx, " WHERE", 0) & " WHERE Client Like'*" & Me.SearchClient & "*' ORDER BY Client, LUnits Desc, LocationName;"
    Me.LocationsList.RowSource = SQLz
End Sub

Private Sub SearchLocation_Button_Click()
    Call ClearSearches("SearchLocation")
    SQLz = SplitString(SQLx, " WHERE", 0) & " WHERE LocationName Like'*" & Me.SearchLocation & "*' ORDER BY LocationName, LUnits Desc;"
    Me.LocationsList.RowSource = SQLz
End Sub

Private Sub SearchShort_Button_Click()
    Call ClearSearches("SearchLShort")
    SQLz = SplitString(SQLx, " WHERE", 0) & " WHERE LocationSN Like'*" & Me.SearchLShort & "*' ORDER BY LocationSN, LUnits Desc;"
    Me.LocationsList.RowSource = SQLz
End Sub

Private Sub SearchCity_Button_Click()
    Call ClearSearches("SearchLCity")
    SQLz = SplitString(SQLx, " WHERE", 0) & " WHERE LCity Like'*" & Me.SearchLCity & "*' ORDER BY LCity, LUnits Desc, LocationName;"
    Me.LocationsList.RowSource = SQLz
End Sub

Private Sub SearchState_Button_Click()
    Call ClearSearches("SearchLState")
    SQLz = SplitString(SQLx, " WHERE", 0) & " WHERE LState Like'*" & Me.SearchLState & "*' ORDER BY LState, LUnits Desc, LocationName;"
    Me.LocationsList.RowSource = SQLz
End Sub

Private Sub SearchCounty_Button_Click()
    Call ClearSearches("SearchLCounty")
    SQLz = SplitString(SQLx, " WHERE", 0) & " WHERE LCounty Like'*" & Me.SearchLCounty & "*' ORDER BY LCounty, LUnits Desc, LocationName;"
    Me.LocationsList.RowSource = SQLz
End Sub

Private Sub SearchZip_Button_Click()
    Call ClearSearches("SearchLZip")
    SQLz = SplitString(SQLx, " WHERE", 0) & " WHERE LZip Like '*" & Me.SearchLZip & "*' Order By LZip, LUnits Desc"
    Me.LocationsList.RowSource = SQLz
End Sub

Private Sub LocationsList_Click()
    Me.clientID = Me.LocationsList.Column(22)
    'On Error GoTo FixMe
    SQLz = "SELECT Locations.*"
    SQLz = SQLz & ", lookup_States.StTaxType"
    SQLz = SQLz & " FROM Locations"
    SQLz = SQLz & " LEFT JOIN lookup_States ON Locations.LState = lookup_States.StateAbbrev"
    SQLz = SQLz & " WHERE LocationId =" & Me.LocationsList.Column(20) & ""
    SQLz = SQLz & " ORDER BY Locations.LocationType, Locations.LState, Locations.LCity;"
    Me.Form.RecordSource = SQLz

    SQLz = "SELECT AssetID"
    SQLz = SQLz & ", Schedule & ' | ' & UnitGroup as SchGrp"
    SQLz = SQLz & ", UnitNum"
    SQLz = SQLz & ", VIN"
    SQLz = SQLz & ", [Status]"
    SQLz = SQLz & ", UnitDescShortVer as UDesc"
    SQLz = SQLz & ", EsitmatedDeliveryDate as EstDelv"
    SQLz = SQLz & ", ActualDeliveryDate as Delv"
    SQLz = SQLz & ", iif(UnitTaxExempt=1, 'Y','N') as Exempt"
    SQLz = SQLz & ", GroupREF as GroupID"
    SQLz = SQLz & ", SchID"
    SQLz = SQLz & " FROM vw_SixKeys As S"
    SQLz = SQLz & " INNER JOIN Units as U On S.AssetId=U.UnitID"
    SQLz = SQLz & " WHERE OffLeaseActual IS NULL AND LocationID=" & Me.LocationsList.Column(20) & " AND MLOrig = 'FA'"
    SQLz = SQLz & " ORDER BY AssetID;"
    Me.UnitsList.RowSource = SQLz

FixMe:
    On Error GoTo 0
End Sub

Private Sub LocationsList_DblClick(Cancel As Integer)
    ' If Me.AssetID & "test" <> "test" Then Call OpenUnit_Click
End Sub

Private Sub ToggleSortAssetID1_Click()
    Call SortListBox("ToggleSortAssetID1", "UnitsList", "AssetID")
End Sub
Public Function SortListBox(CtrlNameX As String, ListBoxX As String, FieldNamex As String)
    Dim Ccrl As Control
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
Public Sub ClearToggles(WhichToggles As Integer)
    Dim Ccrl As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Name Like "*Toggle*" Then
            If WhichToggles = 1 And Me(Ccrl.Name).Name Like "*1" Then Me(Ccrl.Name) = 0
            If WhichToggles = 2 And Me(Ccrl.Name).Name Like "*2" Then Me(Ccrl.Name) = 0
            If WhichToggles = 0 Then Me(Ccrl.Name) = 0
        End If
    Next Ccrl
End Sub

Private Sub ToggleSortClient1_Click()
    Call SortListBox("ToggleSortClient1", "LocationsList", "Client")
End Sub

Private Sub ToggleSortLCity1_Click()
    Call SortListBox("ToggleSortLCity1", "LocationsList", "LCity")
End Sub

Private Sub ToggleSortLCounty1_Click()
    Call SortListBox("ToggleSortLCounty1", "LocationsList", "LCounty")
End Sub

Private Sub ToggleSortLState1_Click()
    Call SortListBox("ToggleSortLState1", "LocationsList", "LState")
End Sub

Private Sub ToggleSortLZip1_Click()
    Call SortListBox("ToggleSortLZip1", "LocationsList", "LZip")
End Sub

Private Sub ToggleSortLocationName1_Click()
    Call SortListBox("ToggleSortLocationName1", "LocationsList", "LocationName")
End Sub

Private Sub ToggleSortShort1_Click()
    Call SortListBox("ToggleSortShort1", "LocationsList", "LocationSN")
End Sub

Private Sub ToggleSortUnitNum1_Click()
    Call SortListBox("ToggleSortUnitNum1", "UnitsList", "UnitNum")
End Sub

Private Sub ToggleSortVIN1_Click()
    Call SortListBox("ToggleSortVIN1", "UnitsList", "VIN")
End Sub

Private Sub UnitsList_DblClick(Cancel As Integer)
    If Me.UnitsList.Column(0) & "test" = "test" Then
        MsgBox "There must be a unit shown first before you can select it", , "No Results"
        Exit Sub
    End If
    Call OpenUnitsForm(Me.UnitsList.Column(0))
    Me.SearchVin.SetFocus
End Sub
