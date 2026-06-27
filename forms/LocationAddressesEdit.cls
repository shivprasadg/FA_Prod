Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim RST        As DAO.Recordset
Dim SQLx, SQLy, SQLz, SourceLocQueryX, SourceUnitQueryX As String
Dim msg1 As String

Private Sub AddNew_Click()
DoCmd.close acForm, "LocationAddressesAdd"
DoCmd.OpenForm "LocationAddressesAdd", , , , , , Me.ClientGroupID
End Sub

Private Sub DeleteMe_Click()
    Dim ResposeX As VbMsgBoxResultEx
    Dim addrx As String
    addrx = Me.LocationName & "<br />" & Me.LCity & ", " & Me.LState
    msg1 = ""
    If Me.LocationID & 0 = 0 Then
        Call Dialog.RichBox("You Must Select an Location before you can delete it", vbCustom + vbInformation, AtlaasVersion, , , 0, 1, 0, 0, "OK")
    End If
    If CInt(UnitsAtAddress([LocationID])) > 0 Then
        msg1 = msg1 & "<b><Font Size=3>Units linked to Address!</font></b><br/>"
        msg1 = msg1 & "You cannot delete Location:<b><br/>" & addrx & "</b><br/>"
        msg1 = msg1 & "<br /><i>Those units would have to be linked<br />to another address beforehand</i><br /> <br /><b>No action taken...</b><br/>"
        ResposeX = Dialog.RichBox(msg1, vbCustom + vbExclamation, AtlaasVersion, , , 0, 1, 0, 0, "OK")
        Exit Sub
    End If
    msg1 = msg1 & "<b><Font Size=3>Delete!</font></b><br/>"
    msg1 = msg1 & "Address:<br /><b>" & addrx & "</b><br/><br/>"
    msg1 = msg1 & "  <b>Are you sure?</b><br/><br/>"
    msg1 = msg1 & "Note:<br/>This will delete this occurrence of this address, only."
    ResposeX = Dialog.RichBox(msg1, vbCustom + vbExclamation, AtlaasVersion, , , 0, 1, 0, 0, "Yes", "No")
    If ResposeX = vbBt2 Then Exit Sub
   ' If ResposeX = vbBt1 Then
        Dim TargetX As String
        TargetX = "TaxStatesBase"
      '  If Me.TaxLevel = "City" Then Targetx = "TaxCitiesBase"
      '  If Me.TaxLevel = "County" Then Targetx = "TaxCountiesBase"
        SQLz = "Delete From Locations WHERE LocationID=" & Me.LocationID & ";"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
        Call RefreshTaxForms
        Me.MoveTo = DMax("LocationID", "LocationAddresses", "ClientGroupId=" & Me.ClientGroupID & "")
        Call MoveRecordTo
    End Sub



Private Sub Form_Current()
Me.LState.Requery
Me.LStateID = Me.LState.Column(4)
  Me.LCounty.RowSource = "SELECT County, CountyID From vw_CCSZ_Lookup WHERE StateID =" & Me.LStateID & " GROUP BY County, CountyID ORDER BY County;"
  Me.LCity.RowSource = "SELECT City, CityID, County, CountyID From vw_CCSZ_Lookup WHERE StateID =" & Me.LStateID & " GROUP BY City, CityID, County, CountyID ORDER BY City, County;"
End Sub

Private Sub Form_Load()
 DoCmd.MoveSize (6 * 1440), , (14 * 1440)
 If Me.OpenArgs & "test" <> "test" Then
    Me.ClientGroupID = CLng(SplitString(Me.OpenArgs, ":", 0))
    Me.MoveTo = CLng(SplitString(Me.OpenArgs, ":", 1))
    Call MoveRecordTo
 End If
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
'SQLz = SQLz & "Select LocationID, LCity, LState, LAddrLine1, FAUnit, Track, Regd, Delv, Titled, LocationName, "
'SQLz = SQLz & "Bill, Gar, HQ, Reg, Title, ClientGroupID "
SQLz = SQLz & "Select * "
SQLz = SQLz & "FROM vw_LocationAddressMatrix "
SQLz = SQLz & "WHERE ClientGroupID=" & Me.ClientGroupID & " "
SQLy = SQLz
SQLz = SQLz & "Order By Lstate, LCity"
Me.LocationAddressList.Form.RecordSource = SQLz
Forms!LocationAddressesEdit!LocationAddressList.SetFocus
Set RST = Forms!LocationAddressesEdit!LocationAddressList.Form.RecordsetClone
RST.FindFirst "LocationID = " & Me.LocationID
Forms!LocationAddressesEdit!LocationAddressList.Form.Bookmark = RST.Bookmark
Set RST = Nothing
'DoCmd.GotoRecord , , acGoTo, Me.CurrentRecord
Me.UnitsAddress.SetFocus
Me.AddressCount = DCount("LocationID", "vw_LocationAddressMatrix", "ClientGroupID=" & Me.ClientGroupID & "")
 DoCmd.MoveSize 8000, , (4.35 * 1440)
End Sub
Private Sub ExpandForm_Click()
If Me.ExpandForm.caption Like "*Expand*" Then
DoCmd.MoveSize , , (14 * 1440)
Me.ExpandForm.caption = "<<Contract"
Else
DoCmd.MoveSize , , (4.35 * 1440)
Me.ExpandForm.caption = "Expand>>"
End If
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

Private Sub LookUpTaxZip_Click()
    If Me.LZip & "test" = "test" Or Me.LState & "test" = "test" Then
        MsgBox "An Address must have both a State and a Zip," & vbNewLine & "in order to perform a 'Look Up Tax (Zip)'", vbCritical, AtlaasVersion
        Exit Sub
    End If
  '  Call LookUpTaxByZipWeb(Me.LState, Me.LPostalCode)
    Call LookUpTaxByZip2Tax(LZip)
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
    Me.LCountry = Me.LState.Column(3)
    Me.LocationTax.Requery
    Dim StateReply As String
    Dim NumRecords As Integer
    Dim Msgx, ResposeX As String
    StateReply = Nz(LookUpByState(Me.LState), "0|")
    If StateReply & "test" = "test" Then StateReply = "0|"
    NumRecords = CInt(SplitString(StateReply, "|", 0))
    Me.LStateID = Me.LState.Column(4)
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

Private Sub LocationFilter_Click()
SQLx = ""
SQLx = SQLx & "Select LocationID, LCity, LState, LAddrLine1, FAUnit, Track, Regd, Delv, Titled, LocationName, "
SQLx = SQLx & "Bill, Gar, HQ, Reg, Title "
SQLx = SQLx & "FROM vw_LocationAddressMatrix "
SQLx = SQLx & "WHERE ClientGroupID=" & Me.ClientGroupID & " "
Dim LocationFilter As String
Select Case LocationFilter
'FAUnit, Track, Regd, Delv, Titled"
'Bill, Gar, HQ, Reg, Title "

Case 1
SQLx = SQLx & "AND FAUnit >0"
Case 2
SQLx = SQLx & "AND Track > 0"
Case 3
SQLx = SQLx & "AND Regd > 0"
Case 4
SQLx = SQLx & "AND Delv > 0"
Case 5
SQLx = SQLx & "AND Titled > 0"
Case 6
SQLx = SQLx & "AND Bill ='Y'"
Case 7
SQLx = SQLx & "AND Gar ='Y'"
Case 8
SQLx = SQLx & "AND HQ ='Y'"
Case 9
SQLx = SQLx & "AND Reg ='Y'"
Case 10
SQLx = SQLx & "AND Title ='Y'"
End Select
 'if Me.SearchState <> "" Then
'If Me.SearchAddress <> "" Then
'If Me.SearchCity <> "" Then SQLx = SQLx & " AND LCity Like '*" & Me.SearchCity & "*'"
'If Me.SearchName <> "" Then SQLx = SQLx & "AND LocationName Like '*" & Me.SearchName & "*'"
SQLx = SQLx & " Order By Lstate, LCity"
Me.LocationAddressList.Form.RecordSource = SQLx
End Sub

Public Sub MoveRecordTo()
    Me.Form.FilterOn = False
    Set RST = Me.RecordsetClone
    RST.FindFirst "LocationID = " & Nz(Me.MoveTo, 0) & ""
    Me.Form.Bookmark = RST.Bookmark
    RST.close
   Set RST = Nothing
End Sub

Private Sub MoveUnits_Click()
DoCmd.OpenForm "MoveLocations"
Forms!MoveLocations.SelectClient = Me.clientID
Forms!MoveLocations.SelectClient_Click
'Forms!MoveLocations.LocationList = Me.LocationID
'Forms!MoveLocations.LocationList_Click
Forms!MoveLocations.SearchAddress1 = Me.LAddrLine1
Forms!MoveLocations.SearchAddress1_Click
End Sub

Private Sub SaveMe_Click()
If Me.Form.Dirty Then Me.Form.Dirty = False
Call RefreshTaxForms
End Sub
