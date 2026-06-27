Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim RST        As DAO.Recordset
Dim SQLx, SQLy, SQLz, SourceLocQueryX, SourceUnitQueryX As String
Dim msg1       As String
Dim StateReply As String
Dim NumRecords As Integer
Dim Msgx, ResposeX As String

Private Sub Form_Current()
'Me.LState.Requery
'Me.LStateID = Me.LState.Column(3)
'  Me.LCounty.RowSource = "SELECT County, CountyID From vw_CCSZ_Lookup WHERE StateID =" & Me.LStateID & " GROUP BY County, CountyID ORDER BY County;"
'  Me.LCity.RowSource = "SELECT City, CityID, County, CountyID From vw_CCSZ_Lookup WHERE StateID =" & Me.LStateID & " GROUP BY City, CityID, County, CountyID ORDER BY City, County;"
End Sub

Private Sub Form_Load()
    Me.ClientGroupID = Nz(Me.OpenArgs, 151)
    'Me.clientID = 3318
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

    ' DoCmd.MoveSize 8000, , (4.35 * 1440)
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


Private Sub SaveMe_Click()
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
    SQLz = SQLz & "," & Nz(Me.clientID, 0) & " "
    SQLz = SQLz & ",'" & Me.LAddrLine1 & "' "
    SQLz = SQLz & ",'" & Me.LAddrLine2 & "' "
    SQLz = SQLz & ",'" & Me.LState.Column(0) & "' "
    SQLz = SQLz & ",'" & Me.LState.Column(2) & "' "
    SQLz = SQLz & ",'" & Me.LCounty & "' "
    SQLz = SQLz & ",'" & Me.LCity & "' "
    SQLz = SQLz & ",'" & Me.LZip & "' "
    SQLz = SQLz & "," & Me.LStateID & " "
    SQLz = SQLz & "," & Me.LCountyID & " "
    SQLz = SQLz & "," & Me.LCityID & " "
    ' SQLz = SQLz & "," & Me.Longitude & " "
    ' SQLz = SQLz & "," & Me.Latitude & " "
    DoCmd.RunSQL (SQLz)
End Sub
