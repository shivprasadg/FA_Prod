Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim RST        As DAO.Recordset
Dim SQLx, SQLy, SQLz, SourceLocQueryX, SourceUnitQueryX, TString As String
Dim msg1       As String
Dim StateReply As String
Dim NumRecords, i As Integer
Dim Msgx, ResposeX As String
Dim stage      As Integer



Private Sub ClientList_Click()
If ClientList.ItemsSelected.Count - 1 < 1 Then TString = "Client" Else TString = "Clients"
Me.SaveMe.caption = "   Connect THIS location to" & vbNewLine & "   the " & ClientList.ItemsSelected.Count & " " & TString & " SELECTED below"
End Sub

Private Sub EditState_Click()
If Me.LocationID & "test" = "test" Then Exit Sub
DoCmd.close acForm, "LocationAddressesEdit"
DoCmd.OpenForm "LocationAddressesEdit", acNormal, , "LocationID=" & Me.LocationID
End Sub

Private Sub Form_Load()
stage = 0
Me.GoNext.Enabled = False
    Me.ClientGroupID = Nz(Me.OpenArgs, 151)
    SQLz = ""
    SQLz = SQLz & "SELECT City, CityID, County, CountyID "
    SQLz = SQLz & "From vw_CCSZ_Lookup "
    SQLz = SQLz & "WHERE StateID =" & Me.LStateID & " "
    SQLz = SQLz & "GROUP BY City, CityID, County, CountyID "
    SQLz = SQLz & "ORDER BY City, County;"
    Me.LCity.RowSource = SQLz

    SQLz = ""
    SQLz = SQLz & "SELECT County, CountyID "
    SQLz = SQLz & "FROM vw_CCSZ_Lookup "
    SQLz = SQLz & "WHERE StateID =" & Me.LStateID & " "
    SQLz = SQLz & "GROUP BY County, CountyID ORDER BY County;"
    Me.LCounty.RowSource = SQLz

    SQLz = ""
    SQLz = SQLz & "Select * "
    SQLz = SQLz & "FROM vw_LocationAddressMatrix "
    SQLz = SQLz & "WHERE ClientGroupID=" & Me.ClientGroupID & " "
    SQLy = SQLz
    SQLz = SQLz & "Order By Lstate, LCity"

    SQLz = ""
    SQLz = SQLz & "SELECT LocationID, LAddressID, ClientShortName, LocationName, LocationSN, ClientGroupID, ClientID "
    SQLz = SQLz & "From LocationNames "
    SQLz = SQLz & "WHERE ClientGroupID=" & Me.ClientGroupID & " Order By LocationName;"
    Me.NamesList.RowSource = SQLz

    SQLz = ""
    SQLz = SQLz & "SELECT clientID, ClientShortName, CompanyName "
    SQLz = SQLz & "From vw_SixKeys "
    SQLz = SQLz & "WHERE ClientGroupID=" & Me.ClientGroupID & " "
    SQLz = SQLz & "Group By clientID, ClientShortName, CompanyName ORDER BY ClientShortName;"
    Me.ClientList.RowSource = SQLz

    SQLz = ""
    SQLz = SQLz & "SELECT LAddressID, LAddrLine1, LState, LCity, LCounty, LPostalCode, LAddrName ,ClientGroupID "
    SQLz = SQLz & "FROM LocationAddresses "
    SQLz = SQLz & "WHERE ClientGroupID =" & Me.ClientGroupID & " ORDER BY LState, LCity;"
    Me.LAddressID.RowSource = SQLz

    If ClientList.ListCount - 1 < 1 Then
        Me.SelectAll.caption = ChrW(9660) & "Select The " & ClientList.ListCount - 1 & " Client"
    Else
        Me.SelectAll.caption = ChrW(9660) & "Select All " & ClientList.ListCount - 1 & " Clients"
    End If
End Sub

Public Sub GoBack_Click()
Select Case stage
    Case 2
        Me.LAddressID = ""
        Me.LocationID = ""
        Me.ClientList = 0
        Me.LocationList = 0
        Me.TabCtrl1.Pages.Item(0).SetFocus
        stage = 0
        Me.GoNext.Enabled = False
        Call ShowStep(1)
    Case 3
        stage = 2
        Call ShowStep(2)
        Me.TabCtrl1.Pages.Item(1).SetFocus
        Me.GoNext.Visible = True
    Case 4
        stage = 3
        Call ShowStep(3)
        Me.TabCtrl1.Pages.Item(1).SetFocus
    End Select

End Sub

Private Sub GoNext_Click()
Select Case stage
    Case 0
        Me.TabCtrl1.Pages.Item(1).SetFocus
    Case 1
        Me.TabCtrl1.Pages.Item(1).SetFocus
        stage = 2
        Call ShowStep(2)
    Case 2
    If ValidateName(2) = False Then Exit Sub
        Me.LocationName2 = Me.LocationName
        Me.Locationsn2 = Me.Locationsn
        Me.TabCtrl1.Pages.Item(2).SetFocus
        stage = 3
        Call ShowStep(3)
        Me.GoNext.Enabled = False
    Case 3
     If ValidateName(3) = False Then Exit Sub
      Me.TabCtrl1.Pages.Item(4).SetFocus
        stage = 4
        Call ShowStep(4)
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

End Sub

Private Sub LocationList_Click()
stage = 1
Me.LAddressID = Me.LocationList.Column(0)
Me.LocationID = Me.LocationList.Column(1)
Call LAddressID_Click
Me.GoNext.Enabled = True
End Sub

Private Sub LZip_AfterUpdate()
    If Len(Me.LZip) < 4 Then Exit Sub
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

Private Sub LState_Click()
    If Nz(Me.LState, "XX") = "XX" Then
        Me.LCity.RowSource = ""
        Me.LCounty.RowSource = ""
        Exit Sub
    End If
    'If Me.LState.Column(1) Like "*RR*" Then Me.LTaxOption = "RR"
    ' Me.LocationTax.Requery
    StateReply = Nz(LookUpByState(Me.LState), "0|")
    NumRecords = CInt(SplitString(StateReply, "|", 0))
    Me.LStateID = Me.LState.Column(3)
    Me.LCounty.RowSource = "SELECT County, CountyID From vw_CCSZ_Lookup WHERE StateID =" & Me.LStateID & " GROUP BY County, CountyID ORDER BY County;"
    Me.LCity.RowSource = "SELECT City, CityID, County, CountyID From vw_CCSZ_Lookup WHERE StateID =" & Me.LStateID & " GROUP BY City, CityID, County, CountyID ORDER BY City, County;"
    If Me.LCity <> "Unknown" And DCount(Nz(Me.LCityID, 0), "vw_CCSZ_Lookup", "StateID=" & Me.LStateID & " AND CityID =" & Nz(Me.LCityID, 0)) = 0 Then Msgx = Msgx & "<br />The City '<b>" & Me.LCity & "</b>' is not within the State of <b> " & Me.LState & "</b>"
    If Me.LCounty <> "Unknown" And DCount(Nz(Me.LCountyID, 0), "vw_CCSZ_Lookup", "StateID=" & Me.LStateID & " AND CountyID =" & Nz(Me.LCountyID, 0)) = 0 Then Msgx = Msgx & "<br />The County '<b>" & Me.LCounty & "</b>' is not within the Sate of <b> " & Me.LState & "</b>"
    If Msgx <> "" Then
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
End Sub


Private Sub NamesList_Click()
Me.LocationName = Me.NamesList.Column(3)
Me.Locationsn = Me.NamesList.Column(4)
End Sub

Private Sub SaveAddress_Click()
  Msgx = ""
    If Me.LAddrLine1 & "test" = "test" Then Msgx = Msgx & HSpaces(3) & "-<b> Street Address 1</b><br />"
    If Me.LCity & "test" = "test" Then Msgx = Msgx & HSpaces(3) & "-<b> City</b><br />"
    If Me.LState & "test" = "test" Then Msgx = Msgx & HSpaces(3) & "-<b> State</b><br />"
    If Me.LCounty & "test" = "test" Then Msgx = Msgx & HSpaces(3) & "-<b> County</b><br />"
    If Msgx <> "" Then
        Msgx = "<Font Size=5><b>Missing Data!</b></Font><br />You're missing the following items:<br />" & Msgx & "<br />Please correct, accordingly..."
        Call Dialog.RichBox(Msgx, vbCustom + vbExclamation, AtlaasVersion, , , 0, 1, 0, 0, "OK")
        Exit Sub
    End If
    Msgx = "<Font Size=4><b>Confrim Address</b></font>"
    Msgx = Msgx & "<br /><br />" & HSpaces(3) & Me.LAddrLine1 & "<br />"
    If Me.LAddrLine2 & "test" <> "test" Then Msgx = Msgx & HSpaces(3) & Me.LAddrLine2 & "<br />"
    Msgx = Msgx & HSpaces(3) & Me.LCity & ", " & Me.LState & ". " & Me.LZip & "<br />"
    Msgx = Msgx & HSpaces(3) & "(" & Me.LCounty & ")<br /><br />"
    Msgx = Msgx & "Do you want to save this address as part of<br /><b>"
    Msgx = Msgx & DLookup("ClientGroupName", "ClientGroups", "ClientGroupID=" & [ClientGroupID]) & "'s</b> lists of addresses?"
    ResposeX = Dialog.RichBox(Msgx, vbCustom + vbInformation, AtlaasVersion, , , 0, 1, 0, 0, "No", "Yes")
    If ResposeX = 101 Then Exit Sub
    SQLz = ""
    SQLz = SQLz & "INSERT INTO LocationAddresses ("
    SQLz = SQLz & "ClientGroupID"
    SQLz = SQLz & ",ClientID"
    SQLz = SQLz & ",LAddrLine1"
    SQLz = SQLz & ",LAddrLine2"
    SQLz = SQLz & ",LCountry"
    SQLz = SQLz & ",LState"
    SQLz = SQLz & ",LCounty"
    SQLz = SQLz & ",LCity"
    SQLz = SQLz & ",LPostalCode"
    SQLz = SQLz & ",LStateID"
    SQLz = SQLz & ",LCountyID"
    SQLz = SQLz & ",LCityID"
    '   SQLz = SQLz & ",Longitude"
    '   SQLz = SQLz & ",Latitude"
    SQLz = SQLz & ") SELECT "
    SQLz = SQLz & " " & Me.ClientGroupID & " "
    SQLz = SQLz & "," & Nz(Me.ClientID, 0) & " "
    SQLz = SQLz & ",'" & Me.LAddrLine1 & "' "
    SQLz = SQLz & ",'" & Me.LAddrLine2 & "' "
    SQLz = SQLz & ",'" & Me.LState.Column(2) & "' "
    SQLz = SQLz & ",'" & Me.LState.Column(0) & "' "
    SQLz = SQLz & ",'" & Me.LCounty & "' "
    SQLz = SQLz & ",'" & Me.LCity & "' "
    SQLz = SQLz & ",'" & Me.LZip & "' "
    SQLz = SQLz & "," & Me.LStateID & " "
    SQLz = SQLz & "," & Me.LCountyID & " "
    SQLz = SQLz & "," & Me.LCityID & " "
    ' SQLz = SQLz & "," & Me.Longitude & " "
    ' SQLz = SQLz & "," & Me.Latitude & " "
   ' DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Me.LocationList.Requery
    Me.LocationList = DMax("LAddressID", "LocationAddresses", "ClientGroupID=" & Me.ClientGroupID & "")
    Me.LAddressID.Requery
    Me.LAddressID = Me.LocationList
End Sub

Private Sub SelectAll_Click()
If ClientList.ListCount - 1 < 1 Then TString = "Client" Else TString = "Clients"
    For i = 0 To ClientList.ListCount - 1
        ClientList.Selected(i) = True
    Next i
Me.SaveMe.caption = "   Connect THIS location to" & vbNewLine & "   the " & ClientList.ListCount - 1 & " " & TString & " SELECTED below"
End Sub
Public Function HowManySelected() As Integer
    i = ClientList.ItemsSelected.Count
HowManySelected = i
End Function


Private Sub LAddressID_Click()
Dim TempString As String
   On Error GoTo FixMe
SQLz = "Select * From LocationAddresses WHERE LAddressID=" & Me.LAddressID.Column(0) & ";"
Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
If Not IsNull(RST!LAddrLine2) Then TempString = TempString & RST!LAddrLine2 & "<br />"
TempString = TempString & RST!LCity & ", " & RST!LState & " " & RST!LPostalCode & "<br />"
TempString = TempString & "<b>County</b> " & RST!LCounty
Me.AddressText = TempString
Me.AddressText2 = Me.LAddressID.Column(1) & "<br />" & TempString
Set RST = Nothing

    SQLz = ""
    SQLz = SQLz & "SELECT LAddressID, LocationID, ClientShortName, LocationName, LocationSN, ClientGroupID, ClientID "
    SQLz = SQLz & "From LocationNames "
    SQLz = SQLz & "WHERE LAddressID=" & Me.LAddressID & " Order By LocationName;"
    Me.AddressUsedList.RowSource = SQLz
Exit Sub
FixMe:
Me.AddressText = ""
Me.AddressText2 = ""
Set RST = Nothing
End Sub
Public Function ValidateName(StageStep As Integer) As Boolean
    ValidateName = True
    Msgx = ""
    If Me.LocationName & "test" = "test" Then Msgx = Msgx & HSpaces(3) & "-<b> Location Name</b><br />"
    If Me.Locationsn & "test" = "test" Then Msgx = Msgx & HSpaces(3) & "-<b> Location Short Name</b><br />"
    If Me.LAddressID & "test" = "test" Then Msgx = Msgx & HSpaces(3) & "-<b> Physical Address</b><br />"
    If StageStep > 2 Then
        If Me.Billing & Me.Garage & Me.Headquarters & Me.Registration & Me.Title & Me.Other & 0 = 0 Then
            Msgx = Msgx & HSpaces(3) & "-<b> Location Type </b><i> (at lease one needs to be selected)</i><br />"
        End If
        If HowManySelected < 1 Then Msgx = Msgx & HSpaces(3) & "-<b> Selected Client </b><i>(at lease one needs to be selected)</i><br />"
    End If
    If Msgx <> "" Then
        Msgx = "<Font Size=5><b>Missing Data!</b></Font><br />You're missing the following items:<br />" & Msgx & "<br />Please correct, accordingly..."
        Call Dialog.RichBox(Msgx, vbCustom + vbExclamation, AtlaasVersion, , , 0, 1, 0, 0, "OK")
        ValidateName = False
    End If
End Function
Private Sub SaveMe_Click()
If ValidateName(3) = False Then Exit Sub
    Msgx = "<Font Size=4><b>Confrim Location</b></font>"
    Msgx = Msgx & "<br /><br />" & "Name:<b> " & Me.LocationName & "</b><br />"
    Msgx = Msgx & "Short:<b> " & Me.Locationsn & "</b><br />"
    If Me.LocationStreet & "test" <> "test" Then Msgx = Msgx & "Company:<b> " & Me.LocationStreet & "</b><br />"
    If Me.AttenToLine & "test" <> "test" Then Msgx = Msgx & "Atten:<b> " & Me.AttenToLine & "</b><br />"
    If Me.LPhone & "test" <> "test" Then Msgx = Msgx & "Ph:<b> " & Me.LPhone & "</b><br />"
    If Me.LFax & "test" <> "test" Then Msgx = Msgx & "Fax:<b> " & Me.LFax & "</b><br />"
    If Me.LEmail & "test" <> "test" Then Msgx = Msgx & "Email:<b>" & Me.LEmail & "</b><br/>"
    Msgx = Msgx & "Address:<br />" & Me.LAddressID.Column(1) & "<br />" & AddressText & "<br /><br />Type:<b>"
    If Me.Billing = True Then Msgx = Msgx & "-Bill"
    If Me.Garage = True Then Msgx = Msgx & " -Gar"
    If Me.Headquarters = True Then Msgx = Msgx & " -HQ"
    If Me.Registration = True Then Msgx = Msgx & " -Reg"
    If Me.Title = True Then Msgx = Msgx & " -Title"
    If Me.Other = True Then Msgx = Msgx & " -Other"
    Msgx = Msgx & "</b><br /><br />Do you want to save and link this Location to <br />"
    If HowManySelected > 1 Then
    Msgx = Msgx & "the<b> " & HowManySelected & "</b> Clients you selected?"""
    Else
    Msgx = Msgx & "<b>" & Me.ClientList.Column(1) & "</b>?"
    End If
    ResposeX = Dialog.RichBox(Msgx, vbCustom + vbInformation, AtlaasVersion, , , 0, 1, 0, 0, "No", "Yes")
    If ResposeX = 101 Then Exit Sub
    Dim ItemIndex As Variant
    For Each ItemIndex In Me.ClientList.ItemsSelected
        SQLz = ""
        SQLz = SQLz & "INSERT INTO LocationNames ("
        SQLz = SQLz & "ClientGroupID"
        SQLz = SQLz & ",ClientID"
        SQLz = SQLz & ",LAddressID"
        SQLz = SQLz & ",LocationName"
        SQLz = SQLz & ",Locationsn"
        SQLz = SQLz & ",LocationStreet"
        SQLz = SQLz & ",AttenToLine"
        SQLz = SQLz & ",LPhone"
        SQLz = SQLz & ",LFax"
        SQLz = SQLz & ",LEmail"
        SQLz = SQLz & ",DivisionID"
        SQLz = SQLz & ",Billing"
        SQLz = SQLz & ",Garage"
        SQLz = SQLz & ",Headquarters"
        SQLz = SQLz & ",Registration"
        SQLz = SQLz & ",Title"
        SQLz = SQLz & ",Other"
        SQLz = SQLz & ") SELECT "
        SQLz = SQLz & " " & Me.ClientGroupID & " "
        SQLz = SQLz & "," & Me.ClientList.ItemData(ItemIndex) & " "
        SQLz = SQLz & "," & Me.LAddressID & " "
        SQLz = SQLz & ",'" & Me.LocationName & "' "
        SQLz = SQLz & ",'" & Me.Locationsn & "' "
        SQLz = SQLz & ",'" & Me.LocationStreet & "' "
        SQLz = SQLz & ",'" & Me.AttenToLine & "' "
        SQLz = SQLz & ",'" & Me.LPhone & "' "
        SQLz = SQLz & ",'" & Me.LFax & "' "
        SQLz = SQLz & ",'" & Me.LEmail & "' "
        SQLz = SQLz & "," & Nz(Me.DivisionID, 0) & " "
        SQLz = SQLz & "," & Nz(Me.Billing, 0) & " "
        SQLz = SQLz & "," & Nz(Me.Garage, 0) & " "
        SQLz = SQLz & "," & Nz(Me.Headquarters, 0) & " "
        SQLz = SQLz & "," & Nz(Me.Registration, 0) & " "
        SQLz = SQLz & "," & Nz(Me.Title, 0) & " "
        SQLz = SQLz & "," & Nz(Me.Other, 0) & " "
 If DCount("LocationName", "LocationNames", "LocationName='" & Me.LocationName & "' AND ClientId=" & Me.ClientList.ItemData(ItemIndex) & "") > 0 Then
   MsgBox "There is already a '" & Me.LocationName & "' location linked to " & DLookup("ClientShortName", "vw_SixKeys", "ClientId=" & Me.ClientList.ItemData(ItemIndex) & "") & " " & vbNewLine & "  (location was not added)", vbCritical, AtlaasVersion
  Else
         DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
 End If
        Next
    End Sub


Public Function SortListBox(CtrlNameX As String, ListBoxX As String, FieldNamex As String)
    Dim Ccrl   As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Name Like "*Toggle*" And Me(Ccrl.Name).Name <> CtrlNameX Then Me(Ccrl.Name) = 0
        '    If CtrlNameX Like "*1" And Me(Ccrl.Name).Name Like "*1" Then Me(Ccrl.Name) = 0
        '    If CtrlNameX Like "*2" And Me(Ccrl.Name).Name Like "*2" Then Me(Ccrl.Name) = 0
        'End If
    Next Ccrl
    SQLz = SplitString(Me(ListBoxX).RowSource, " Order By", 0) & " Order By " & FieldNamex & ""
    If Me(CtrlNameX) = 0 Then SQLz = SQLz & " Desc"
    Me(ListBoxX).RowSource = SQLz
End Function
Private Sub ToggleSortAddress_Click()
Call SortListBox("ToggleSortAddress", "LocationList", "LocationAddresses.LAddrLine1")
End Sub
Private Sub ToggleSortState_Click()
Call SortListBox("ToggleSortState", "LocationList", "LocationAddresses.LState")
End Sub
Private Sub ToggleSortCity_Click()
Call SortListBox("ToggleSortCity", "LocationList", "LocationAddresses.LCity")
End Sub
Private Sub ToggleSortCounty_Click()
Call SortListBox("ToggleSortCounty", "LocationList", "LocationAddresses.LCounty")
End Sub
Private Sub ToggleSortType_Click()
Call SortListBox("ToggleSortType", "LocationList", "IIf([bill]='Y','-Bill',Null) & IIf([Gar]='Y','-Gar',Null) & IIf([HQ]='Y','-HQ',Null) & IIf([Reg]='Y','-Reg',Null) & IIf(vw_LocationAddressMatrix.[Title]='Y','-Til',Null)")
End Sub
Private Sub ClearSorts_Click()
ToggleSortAddress = 0
ToggleSortState = 1
ToggleSortCity = 0
ToggleSortCounty = 0
ToggleSortType = 0
Me.LocationList.RowSource = SplitString(Me.LocationList.RowSource, "Order By", 0) & " Order By LocationAddresses.LState"
End Sub
