Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim RST        As DAO.Recordset
Dim SQLz, SourceLocQueryX, SourceUnitQueryX, XLS, Headerx As String
Dim PrevBilling As String

Private Sub AddNewDivision_Click()
    DoCmd.close acForm, "AddClientDivision"
    DoCmd.OpenForm "AddClientDivision"
    Forms!AddClientDivision.Form.ClientID = Me.Parent.ClientID
    Forms!AddClientDivision.Form.ClientGroupID = Me.Parent.ClientGroupID
End Sub

Private Sub ContactsCBO_Click()
    Dim Msgx, ClassificationX As String
    Dim LinkTypeX As Integer
    Msgx = ""
    If IsNull(Me.LocationID) Then Msgx = Msgx & vbNewLine & "   Location"
    'If IsNull(Me.ContactsCBO.Column(0)) Then MsgX = MsgX & vbNewLine & "   Contact to Link"
    'If IsNull(Me.LinkTypeCBO.Column(0)) Then MsgX = MsgX & vbNewLine & "   Type of Link to use"
    If Msgx <> "" Then
        Msgx = "You must Select a" & Msgx
        MsgBox Msgx, vbCritical, "Missing Data"
        Exit Sub
    End If
    ClassificationX = DLookup("contactClassification", "Contacts", "contactID=" & Me.ContactsCBO.Column(0))
    LinkTypeX = DLookup("LinkTypeID", "ContactsLinkedToLocationTypes", "LinkType='" & ClassificationX & "'")
    If DLookup("ContactID", "ContactsLinkedToLocation", "ContactID=" & Me.ContactsCBO.Column(0) & " AND LinkTypeID =" & LinkTypeX & " AND LocationID=" & Me.LocationID) Then
        MsgBox Me.ContactsCBO.Column(2) & vbNewLine & "Is already linked  to " & Me.LocationName & vbNewLine & "as a '" & ClassificationX & "' contact type" & vbNewLine & vbNewLine & "Canceled...", vbCritical, "We've Done this already"

        Exit Sub
    End If
    SQLz = "INSERT INTO ContactsLinkedToLocation ( ContactID, LocationID, LinkTypeID ) " & _
        "SELECT " & Me.ContactsCBO.Column(0) & ", " & Me.LocationID & ", " & LinkTypeX & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Me.LinkedContactsList.Requery
    Me.Parent.ClientContact_sub.Form.LinkedLocationsList.Requery
    Call SetUnits
    Me.ContactDetail = ""
End Sub

Private Sub DeleteDivision_Click()
    Dim UnitCnt As Integer
    UnitCnt = DCount("UnitID", "Units", "DivisionOwnerID=" & Me.DivisionList.Column(0))
    If UnitCnt > 0 Then
        MsgBox "You Cannot Delete Division:" & vbNewLine & " '" & Me.DivisionList.Column(1) & "'" & vbNewLine & "There are  " & UnitCnt & " units in that Division", vbCritical, "Don't Run with Matches!"
        Exit Sub
    End If
    Dim Askit   As String
    Askit = MsgBox("Are you sure you want DELETE Division:" & vbNewLine & Me.DivisionList.Column(1) & "?", vbCritical + vbYesNo, "Don't play with matches")
    If Askit <> "6" Then Exit Sub
    SQLz = "Delete From ClientDivisions WHERE DivisionID =" & Me.DivisionList.Column(0) & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Me.DivisionList.Requery
    Me.DivisionID.Requery
End Sub

Private Sub DeleteLocation_Click()
    If FAUserAccess < 75 Then Exit Sub

    Dim UnitCnt As Integer
    UnitCnt = DCount("UnitID", "Units", "UnitGarageLocationRef=" & Me.LocationList.Column(0))
    UnitCnt = UnitCnt + DCount("UnitID", "Units", "UnitDeliveryLocationRef=" & Me.LocationList.Column(0))
    UnitCnt = UnitCnt + DCount("UnitID", "Units", "UnitTitleLocationRef=" & Me.LocationList.Column(0))
    UnitCnt = UnitCnt + DCount("UnitID", "Units", "UnitRegLocationRef=" & Me.LocationList.Column(0))
    ' UnitCnt = UnitCnt + DCount("UnitID", "Units", "UCCC=" & Me.LocationList.Column(0))
    If UnitCnt > 0 Then
        MsgBox "You Cannot Delete " & Me.LocationList.Column(1) & " Location:" & vbNewLine & " '" & Me.LocationList.Column(2) & "'" & vbNewLine & "There are  " & UnitCnt & " units at that Location", vbCritical, "Don't Run with Matches!"
        Exit Sub
    End If

    Dim Askit   As String
    Askit = MsgBox("Are you sure you want DELETE " & Me.LocationList.Column(1) & " Location:" & vbNewLine & Me.LocationList.Column(2) & "?", vbCritical + vbYesNo, "Don't play with FIRE")
    If Askit <> "6" Then Exit Sub
    SQLz = "Delete From Locations WHERE LocationID =" & Me.LocationList.Column(0) & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Me.LocationList.Requery
End Sub

Private Sub DivisionID_Click()
    Call SyncUnitsToLocationDivisions
    MsgBox "Locations / Divisions Synced to the Divisions Assigned to the Units' Locations" & vbNewLine & "(if any needed to change)", vbInformation
    Me.DivisionID.Requery
End Sub

Private Sub ExportLocations_Click()

   SQLz = "Select LocationType as Type"
   SQLz = SQLz & " ,LocationName"
   SQLz = SQLz & " ,Locationsn"
   SQLz = SQLz & " ,LocationStreet"
   SQLz = SQLz & " ,LAddrLine1"
   SQLz = SQLz & " ,LAddrLine2"
   SQLz = SQLz & " ,LAddrLine3"
   SQLz = SQLz & " ,LCity"
   SQLz = SQLz & " ,LCounty"
   SQLz = SQLz & " ,LState"
   SQLz = SQLz & " ,LZip"
   SQLz = SQLz & " ,LCountry"
   SQLz = SQLz & " ,LTaxOption As [Tax Option]"
   SQLz = SQLz & " ,Locations.LastTaxUpdate AS [Last Tax Update]"
   SQLz = SQLz & " ,LTaxRate As [Loc Tax Rate]"
   SQLz = SQLz & " ,StateSales AS State"
   SQLz = SQLz & " ,CountySales AS County"
   SQLz = SQLz & " ,CitySales AS City"
   SQLz = SQLz & " ,Nz(LTaxRate,0) - (Nz(StateSales,0) + Nz(CountySales,0) + Nz(CitySales,0)) AS [Tax Var]"
   SQLz = SQLz & " ,Locations.LocationID"
   SQLz = SQLz & " ,DivisionID,Locations.LStateID"
   SQLz = SQLz & " ,Locations.LCountyID"
   SQLz = SQLz & " ,Locations.LCityID"
   SQLz = SQLz & " ,ClientGroupID"
   SQLz = SQLz & " From Locations"
   SQLz = SQLz & " LEFT JOIN vw_LocationTaxes AS T ON Locations.LocationID = T.LocationID"
   SQLz = SQLz & " WHERE ClientGroupID =" & Me.ClientGroupID & " "
   SQLz = SQLz & " ORDER BY LocationType, LocationName"

   XLS = "|A,mTable:Name=Locations|A,DD|"
   XLS = XLS & "A,W14|A,TC|J,BC|L,TC|M,BC|M,W11|O,BC|"
   XLS = XLS & "O,cColor:3:green|P,cColor:3:crimson:End=S|S,cColor:3:Orange:End=S|"
   XLS = XLS & "O,cBorder:Med:End=O|P,cBorder:Med:End=S|S,cBorder:Med:End=S|"
   XLS = XLS & "O,P3|P,P3|Q,P3|R,P3|S,P3|N,W11|O,W08|"
   XLS = XLS & "O,TC|P,TC|Q,TC|R,TC|S,TC|T,TC|"
   XLS = XLS & "N,cLabel:End=S:2:" & TLeftLine & LineSpace(18) & " Location Taxes " & LineSpace(20) & TRightLine & "|"
   XLS = XLS & "T,HH|U,HH|V,HH|W,HH|X,HH|Y,HH|"
    Headerx = "Location Export - " & Me.Parent.ClientShNm & " - " & Format(Date, "mm-dd-yy")
    MessageUser "Open", Headerx, "Please wait this will take a moment..."
    SendReport2Excel SQLz, "Locations", Headerx, True, XLS, 0, False
    MessageUser "Close"
End Sub

Private Sub LAddrLine1_AfterUpdate()
    Call BillingAddrChange
End Sub

Private Sub LAddrLine1_GotFocus()
   PrevBilling = Nz(Me.LAddrLine1, "")
End Sub

Private Sub LAddrLine2_AfterUpdate()
    Call BillingAddrChange
End Sub

Private Sub LAddrLine2_GotFocus()
    PrevBilling = Nz(Me.LAddrLine1, "")
End Sub

Private Sub LinkedContactsList_Click()
    SQLz = "Select * From Contacts WHERE ContactID =" & Me.LinkedContactsList.Column(0) & ";"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    Me.ContactDetail = RST![contactFullName] & vbNewLine & RST!contactTitle & vbNewLine & _
        Me.LAddrLine1 & vbNewLine & _
        Me.[LCity] & " " & Me.LState & " " & Me.LZip & vbNewLine & _
        "Phone: " & Left(RST![contactPrimaryPhone], 3) & "-" & Mid(RST![contactPrimaryPhone], 4, 3) & "-" & Right(RST![contactPrimaryPhone], 4) & vbNewLine & _
        "Email: " & RST![contactEmail] & vbNewLine
    RST.close
    Set RST = Nothing
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
    Me.LocationList.Requery
End Sub

Private Sub Location_Add2_AfterUpdate()
    Call BillingAddrChange
End Sub
Private Sub BillingAddrChange()

  If PrevBilling & "test" <> "test" And PrevBilling <> Me.LAddrLine1 And Me.Location_Type = "Billing" Then
            Dim msg1 As String
            msg1 = msg1 & "<html><body><p style='color:RED;font-family:Segoe UI;font-size:20px;'><b>Billing Location Change</b></p>"
            msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:12px'>To: Synd Dept/Accounting<br />"
            msg1 = msg1 & "Subject:&#160;&#160;&#160;&#160;&#160;&#160;Billing Location Changed, Re Client Billing for:<b> " & Me.Parent.clientCompanyName & "</b><br />"
            msg1 = msg1 & "Priority:<b>&#160;&#160;&#160;&#160;&#160;&#160;High </b><br/><br/>"
            msg1 = msg1 & "All,<br/><br/>There was a change submitted to a Billing Location, the details are below.<br /><br />"
            msg1 = msg1 & "Client: <B>" & Me.Parent.clientCompanyName & "</B><br />"
            msg1 = msg1 & "Prev Billing Addr: <B>" & PrevBilling & "</B><br />"
            msg1 = msg1 & "New Billin Addr:<br /> <B>"
            msg1 = msg1 & " " & Me.LocationName & "<br />"
            msg1 = msg1 & " " & Me.LocationStreet & "<br />"
            msg1 = msg1 & " " & Me.LAddrLine1 & "<br />"
            msg1 = msg1 & " " & Me.LCity & ", " & Me.LState & ". " & Me.LZip & "</B><br />"
            msg1 = msg1 & "<br />"
            msg1 = msg1 & "User who made change: <B>" & FAUserFullName & "</B><br />"
            msg1 = msg1 & "<p style='font-family:Arial ; font-size:11px ; margin-left:30px'></P>"
            msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:8px'><br/>[End Msg]"
            msg1 = msg1 & "Email Created: " & Date & " (via ATLAAS Core Data Manager).<br/><br />"
            msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:12px'><br/>"
            msg1 = msg1 & "Regards,<br /><br /><br />"
            msg1 = msg1 & "</body></html>"
     '      Call SendEmailMessage("ljohnson@fleetadvantage.com", "Billing Location Change", msg1, True, True, False, "")
            Call SendEmailMessage("faaccounting@fleetadvantage.com", "Term Penatly Change", msg1, True, True, False, "closings@fleetadvantage.com")
   End If
End Sub

Private Sub LocationName_Click()
    If Me.LocationStreet & "test" = "test" Then Me.LocationStreet = Me.Parent.clientCompanyName
End Sub

Private Sub Locationsn_AfterUpdate()
    If Me.LocationStreet & "test" = "test" Then Me.LocationStreet = Me.Parent.clientCompanyName
End Sub

Private Sub LocationsReportPrintView_Click()
    DoCmd.close acReport, "LocationDirectory"
    DoCmd.OpenReport "LocationDirectory", acViewPreview
End Sub

Private Sub LocationsReportView_Click()
    DoCmd.close acReport, "LocationDirectory"
    DoCmd.OpenReport "LocationDirectory", acViewReport
End Sub

Private Sub LookUpTaxZip_Click()
    If Me.LState & "test" = "test" Or Me.LZip & "test" = "test" Then
        MsgBox "An Address must have both a State and a Zip," & vbNewLine & "in order to perform a 'Look Up Tax (Zip)'", vbCritical, AtlaasVersion
        Exit Sub
    End If
    Call LookUpTaxByZip2Tax(LZip)
End Sub

Private Sub LState_Click()
    If Nz(Me.LState, "XX") = "XX" Then
        Me.LCity.RowSource = ""
        Me.LCounty.RowSource = ""
        Exit Sub
    End If

    If Me.LState.Column(1) Like "*RR*" Then Me.LTaxOption = "RR"
    Me.LocationTax.Requery
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

    'If Me.LCity <> "Unknown" And DCount(Nz(Me.LCityID, 0), "vw_CCSZ_Lookup", "StateID=" & Me.LStateID & " AND CityID =" & Nz(Me.LCityID, 0)) = 0 Then Msgx = Msgx & "<br />The City '<b>" & Me.LCity & "</b>' is not within the State of <b> " & Me.LState & "</b>"
    'If Me.LCounty <> "Unknown" And DCount(Nz(Me.LCountyID, 0), "vw_CCSZ_Lookup", "StateID=" & Me.LStateID & " AND CountyID =" & Nz(Me.LCountyID, 0)) = 0 Then Msgx = Msgx & "<br />The County '<b>" & Me.LCounty & "</b>' is not within the Sate of <b> " & Me.LState & "</b>"
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
    ' Me.LCounty = CheckCounty(Me.LCounty, Me.LState, Me.LCity)
    ' Me.LCity = CheckCity(Me.LCounty, Me.LState, Me.LCity)
    If NumRecords = 0 Then
        Me.LCity.RowSource = ""
        Me.LCounty.RowSource = ""
        Exit Sub
    End If
    If Me.Form.Dirty Then Me.Form.Dirty = False
End Sub

Private Sub LTaxRate_AfterUpdate()
       Me.LastTaxUpdate = Date
       Me.LocationTax.Requery
End Sub

Private Sub LZip_AfterUpdate()
    If Len(Me.LZip) Or IsNull(Me.LZip) < 4 Then Exit Sub
    Dim ZipReply As String
    Dim NumRecords As Integer
    ZipReply = LookUpByZip(Left(Me.LZip, 5))
    If ZipReply = "" Then NumRecords = 0 Else NumRecords = CInt(SplitString(ZipReply, "|", 0))
    If NumRecords = 0 Then
        LookUpZipWeb (Left(Me.LZip, 5))
        Exit Sub
    End If
    Me.LCity.RowSource = SplitString(ZipReply, "|", 4)
    Me.LCounty.RowSource = SplitString(ZipReply, "|", 5)
    If NumRecords > 1 Then
        MsgBox " The Zip you entered has more than one record." & vbNewLine & "Please make sure to select the proper City and County"
    End If
    If Nz(Me.LCity, "Unknown") = "Unknown" Then Me.LCity = SplitString(ZipReply, "|", 1)
    If Nz(Me.LState, "Unknown") = "Unknown" Then Me.LState = SplitString(ZipReply, "|", 2)
    If Nz(Me.LCounty, "Unknown") = "Unknown" Then Me.LCounty = SplitString(ZipReply, "|", 3)
    'If Rst!City <> Me.LCity Then
    '    MsgBox "The Zip you entered returned a State (" & Rst!State & ")that is different then the one entered (" & Me.LState & ")"
    'End If
    'If Rst!State <> Me.LState Then

    'End If
    'If Rst!county <> Me.LCounty Then

    'End If
    Me.LCity = SplitString(ZipReply, "|", 1)
    Me.LState = SplitString(ZipReply, "|", 2)
    Me.LCounty = SplitString(ZipReply, "|", 3)

    ' Dim ZipColl As Object
    ' Set ZipColl = New Collection

    '  Set ZipColl = LookUpZipWeb(Left(Me.LZip, 5))
    'If ZipColl("AliasCities") > 0 Then
    'Dim Msx As String
    'Dim ii As Integer
    'For ii = 0 To CInt(ZipColl("AliasCities"))
    'Msx = Msx & ZipColl("AlaisCity" & (ii))
    'Next ii
    'MsgBox Msx
    'End If
End Sub

Private Sub SaveMe_Click()
    If Me.Dirty Then Me.Dirty = False
    Me.Form.Requery
    'Me.LocationList.Requery
End Sub

Private Sub SyncUnitsLocationDivisions_Click()
    Call SyncUnitsToLocationDivisions
    MsgBox "Locations / Divisions Synced to the Divisions Assigned to the Units' Locations" & vbNewLine & "(if any needed to change)", vbInformation
End Sub

Private Sub ToggleSortName_Click()
    Call SortListBox("ToggleSortName", "LocationList", "LocationName")
End Sub
Private Sub ToggleSortState_Click()
    Call SortListBox("ToggleSortState", "LocationList", "LState")
End Sub
Private Sub ToggleSortTax_Click()
    Call SortListBox("ToggleSortTax", "LocationList", "LTaxRate")
End Sub
Private Sub ToggleSortType_Click()
    Call SortListBox("ToggleSortType", "LocationList", "LocationType")
End Sub
Private Sub ToggleSortUnits_Click()
    Call SortListBox("ToggleSortUnits", "LocationList", "Count(UnitId)")
End Sub
Public Function SortListBox(CtrlNameX As String, ListBoxX As String, FieldNamex As String)
    Dim Ccrl   As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Name Like "*Toggle*" And Me(Ccrl.Name).Name <> CtrlNameX Then Me(Ccrl.Name) = 0
    Next Ccrl
    SQLz = SplitString(Me(ListBoxX).RowSource, " Order By", 0) & " Order By " & FieldNamex & ""
    If Me(CtrlNameX) = 0 Then SQLz = SQLz & " Desc"
    Me(ListBoxX).RowSource = SQLz
End Function
Private Sub ClearSorts_Click()
    ToggleSortName = 1
    ToggleSortState = 0
    ToggleSortTax = 0
    ToggleSortType = 0
    ToggleSortUnits = 0
    Me.LocationList.RowSource = SplitString(Me.LocationList.RowSource, "Order By", 0) & " Order By LocationName"
End Sub
Private Sub UnLinkedContact_Click()
    If IsNull(Me.LinkedContactsList.Column(0)) Then Exit Sub
    Dim Askit  As String
    Askit = MsgBox("You're about to Un-Link..." & vbNewLine & "Name: " & Me.LinkedContactsList.Column(3) & vbNewLine & "From:  " & Me.LocationName & vbNewLine & "Type:   " & Me.LinkedContactsList.Column(2) & vbNewLine & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Un-Link")
    If Askit <> 6 Then Exit Sub
    SQLz = "Delete From ContactsLinkedToLocation " & _
        "WHERE  ContactID=" & Me.LinkedContactsList.Column(0) & " AND LinkTypeID=" & Me.LinkedContactsList.Column(5) & " AND LocationID=" & Me.LocationID & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Me.LinkedContactsList.Requery
    Me.Parent.ClientContact_sub.Form.LinkedLocationsList.Requery
    Call SetUnits
    Me.ContactDetail = ""
End Sub
Private Sub AddNew_Click()
    'On Error Resume Next
    If Me.Dirty Then Me.Dirty = False
    On Error GoTo FixMe
    DoCmd.GotoRecord , , acNewRec
    Exit Sub
FixMe:
    DoCmd.GotoRecord , , acNewRec
    On Error GoTo 0
End Sub

Private Sub Form_Close()
    On Error GoTo Form_Close_Error
    On Error GoTo 0
    Exit Sub
Form_Close_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Close of VBA Document Form_Locations_sub"
End Sub

Public Sub Form_Load()
    SourceUnitQueryX = "vw_SixKeys"
    Me.ContactDetail = ""
End Sub
Public Sub Form_Current()
    'Call SetLocations
End Sub
Public Sub SetLocations()
    Me.SearchAssetID = ""
    Me.SearchUnit = ""
    Me.SearchLease = ""
    Me.SearchSchedule = ""
    Me.SearchDesc = ""
    Me.SearchStatus = ""
    Me.LocationList.RowSource = ""
    Me.LocationList.Requery
    On Error GoTo Form_Current_Error

    Dim ClientX, ClientGroupx As Long
    ClientX = Nz(Me.Parent.ClientID, 0)
    ClientGroupx = Nz(Me.Parent.ClientGroupID, 0)

    SQLz = "SELECT DivisionID, DivisionName, ClientID From ClientDivisions " & _
        "WHERE Deleted=False and ClientGroupID=" & ClientGroupx & " Order by DivisionName"
    Me.DivisionID.RowSource = SQLz

    SQLz = "SELECT DivisionID, DivisionName, ClientID, ClientGroupID From ClientDivisions " & _
        "WHERE Deleted=False and ClientID=" & ClientX & " Order by DivisionName"
    Me.DivisionList.RowSource = SQLz

    SQLz = "SELECT L.LocationID, L.LocationType AS Type, L.LocationName AS Name, L.LState AS State, Count(U.UnitID) AS Units"
    SQLz = SQLz & ", Format([LTaxRate],'Percent') AS TaxRate, L.LocObsolete, L.ClientREF "
    SQLz = SQLz & "FROM Locations AS L LEFT JOIN Units AS U ON L.LocationID = U.UnitGarageLocationRef "
    SQLz = SQLz & "WHERE L.ClientGroupID =" & ClientGroupx & " "
    SQLz = SQLz & "GROUP BY L.LocationID, L.LocationType, L.LocationName, L.LState, L.LTaxRate, L.LocObsolete, L.ClientREF "
    SQLz = SQLz & "ORDER BY L.LocationType, L.LocationName, L.LocObsolete;"

    Me.LocationList.RowSource = SQLz
    On Error Resume Next
    Me.LocationList.Selected(0) = True
    On Error GoTo Form_Current_Error
    Call LocationList_Click
    Me.ContactDetail = ""
    Exit Sub
Form_Current_Error:
    Dim m      As String
    If DCount("LocationID", "Locations", "ClientGroupID=" & Me.Parent.ClientGroupID & "") = 0 Then
        SQLz = "Insert into Locations (ClientGroupID, LocationType, LocationName, Locationsn, LocationStreet) Select " & Me.Parent.ClientGroupID & ",'Garage','TBD','TBD','" & Me.Parent.clientCompanyName & "';"
        m = "Error! But Don't Panic, Atlaas Knows what to do..." & vbNewLine
        m = m & "There isn't a 'Location' assigend to Client:" & Me.Parent.clientCompanyName & vbNewLine
        m = m & "Atlaas Core will assign a location named 'TBD' for you." & vbNewLine & vbNewLine & "Please Press 'OK'"
        MsgBox m, vbInformation, "Atlaas Loves You"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
        Call SetLocations
    End If
    Exit Sub
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Current of VBA Docume"
End Sub

Private Sub SetUnits()
    Me.SearchAssetID = ""
    Me.SearchUnit = ""
    Me.SearchLease = ""
    Me.SearchSchedule = ""
    Me.SearchDesc = ""
    Me.SearchStatus = ""

    SQLz = "Select AssetID, MLNO, Schedule, UnitNum, UnitDescShortVer as Descp, Status From vw_SixKeys " & _
        "Where LocationID=" & Me.LocationID & " Order by MLOrig, AssetID;"
    Me.UnitsList.RowSource = SQLz

    Dim ListCountX As String
    ListCountX = "0 Units @ Location"
    If Me.UnitsList.ListCount > 0 Then ListCountX = Me.UnitsList.ListCount - 1 & " Units @ Location"
    If (Me.UnitsList.ListCount - 1) = 1 Then ListCountX = Me.UnitsList.ListCount - 1 & " Unit @ Location"
    Me.UNITS.caption = ListCountX
    SQLz = "SELECT ContactID, LocationID, ContactType, [FirstName] & ' ' & [LastName] AS Name, Email, Phone From vw_LocationContacts Where LocationID= " & Me.LocationID & ";"
    Me.LinkedContactsList.RowSource = SQLz
End Sub

Public Sub LocationList_Click()
    'On Error GoTo Fixme
    Dim LocX   As Integer
    LocX = Me.LocationList.Column(0)
    SQLz = "SELECT Locations.*, lookup_States.StTaxType " & _
        "FROM Locations LEFT JOIN lookup_States ON Locations.LState = lookup_States.StateAbbrev " & _
        "WHERE Locations.LocationID = " & LocX & ";"
    Me.Form.RecordSource = SQLz

    Me.LCounty.RowSource = "SELECT County, CountyID From vw_CCSZ_Lookup WHERE StateID =" & Me.LStateID & " GROUP BY County, CountyID ORDER BY County;"
    Me.LCity.RowSource = "SELECT City, CityID, County, CountyID From vw_CCSZ_Lookup WHERE StateID =" & Me.LStateID & " GROUP BY City, CityID, County, CountyID ORDER BY City, County;"
    Call SetUnits
    Me.ContactDetail = ""
    Me.LocationList = LocX
FixMe:
End Sub



Private Sub SearchAssetId_Change()
    'Me.SearchAssetId = ""
    Me.SearchUnit = ""
    Me.SearchLease = ""
    Me.SearchSchedule = ""
    Me.SearchDesc = ""
    Me.SearchStatus = ""
    SQLz = "Select AssetID, MLNO, Schedule, UnitNum, Desc, Status From " & SourceUnitQueryX & " Where GarageLocationID=" & Me.LocationList.Column(0) & " " & _
        "AND AssetID Like '*" & Me.SearchAssetID.Text & "*' Order by AssetID;"
    Me.UnitsList.RowSource = SQLz
End Sub
Private Sub SearchLease_Change()
    Me.SearchAssetID = ""
    Me.SearchUnit = ""
    'Me.SearchLease = ""
    Me.SearchSchedule = ""
    Me.SearchDesc = ""
    Me.SearchStatus = ""
    SQLz = "Select AssetID, MLNO, Schedule, UnitNum, Desc, Status From " & SourceUnitQueryX & " Where GarageLocationID=" & Me.LocationList.Column(0) & " " & _
        "AND MLNo Like '*" & Me.SearchLease.Text & "*' Order by AssetID;"
    Me.UnitsList.RowSource = SQLz
End Sub
Private Sub SearchSchedule_Change()
    Me.SearchAssetID = ""
    Me.SearchUnit = ""
    Me.SearchLease = ""
    'Me.SearchSchedule = ""
    Me.SearchDesc = ""
    Me.SearchStatus = ""
    SQLz = "Select AssetID, MLNO, Schedule, UnitNum, Desc, Status From " & SourceUnitQueryX & " Where GarageLocationID=" & Me.LocationList.Column(0) & " " & _
        "AND Schedule Like '*" & Me.SearchSchedule.Text & "*' Order by AssetID;"
    Me.UnitsList.RowSource = SQLz
End Sub
Private Sub SearchUnit_Change()
    Me.SearchAssetID = ""
    'Me.SearchUnit = ""
    Me.SearchLease = ""
    Me.SearchSchedule = ""
    Me.SearchDesc = ""
    Me.SearchStatus = ""
    SQLz = "Select AssetID, MLNO, Schedule, UnitNum, Desc, Status From " & SourceUnitQueryX & " Where GarageLocationID=" & Me.LocationList.Column(0) & " " & _
        "AND UnitNum Like '*" & Me.SearchUnit.Text & "*' Order by UnitNum;"
    Me.UnitsList.RowSource = SQLz
End Sub
Private Sub SearchDesc_Change()
    Me.SearchAssetID = ""
    Me.SearchUnit = ""
    Me.SearchLease = ""
    Me.SearchSchedule = ""
    'Me.SearchDesc = ""
    Me.SearchStatus = ""
    SQLz = "Select AssetID, MLNO, Schedule, UnitNum, Desc, Status From " & SourceUnitQueryX & " Where GarageLocationID=" & Me.LocationList.Column(0) & " " & _
        "AND [Desc] Like '*" & Me.SearchDesc.Text & "*' Order by [Desc],AssetID;"
    Me.UnitsList.RowSource = SQLz
End Sub
Private Sub SearchStatus_Change()
    Me.SearchAssetID = ""
    Me.SearchUnit = ""
    Me.SearchLease = ""
    Me.SearchSchedule = ""
    Me.SearchDesc = ""
    'Me.SearchStatus = ""
    SQLz = "Select AssetID, MLNO, Schedule, UnitNum, Desc, Status From " & SourceUnitQueryX & " Where GarageLocationID=" & Me.LocationID & " " & _
        "AND [Status] Like '*" & Me.SearchStatus.Text & "*' Order by [Desc],AssetID;"
    Me.UnitsList.RowSource = SQLz
End Sub
Private Sub UnitsList_DblClick(Cancel As Integer)
    If Me.UnitsList.Column(0) & "test" = "test" Then
        MsgBox "There must be a unit shown first before you can select it", , "No Results"
        Exit Sub
    End If
    Call OpenUnitsForm(Me.UnitsList.Column(0))
    'DoCmd.OpenForm "UnitsForm", , , "UNITID=" & Me.UnitsList.Column(0) & ""
    Me.LocationList.SetFocus
End Sub
