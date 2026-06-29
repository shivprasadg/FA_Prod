Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim RST As DAO.Recordset
Dim SQLz As String


Private Sub AddNew_Click()
DoCmd.GotoRecord , , acNewRec
End Sub

Private Sub Form_Current()
Me.LocationList.Requery
Me.LocationList = Me.LocationID
Call LocationList_Click
Dim ClientX
ClientX = Nz(Me.ClientID, 0)
    Me.SearchAssetID = ""
    Me.SearchUnit = ""
    Me.SearchLease = ""
    Me.SearchSchedule = ""
SQLz = "SELECT DivisionID, DivisionName, ClientID From ClientDivisions " & _
       "WHERE Deleted=False and ClientID=" & ClientX & " Order by DivisionName"
Me.DivisionID.RowSource = SQLz
SQLz = "SELECT LocationID, LocationType AS Type, LocationName AS Name, LState AS State, Units, LocObsolete, ClientREF " & _
       "FROM ParkedLocationsUnitCount WHERE ClientREF=" & ClientX & " ORDER BY LocationType, LocObsolete, LocationName;"
Me.LocationList.RowSource = SQLz
End Sub

Private Sub Form_Load()
SQLz = "SELECT LocationID, LocationType AS Type, LocationName AS Name, LState AS State, Units, LocObsolete, ClientREF " & _
"FROM ParkedLocationsUnitCount " & _
"WHERE ClientREF=" & Me.OpenArgs & " ORDER BY LocationType, LocObsolete, LocationName;"
Me.LocationList.RowSource = SQLz
End Sub

Private Sub LocationList_Click()
On Error GoTo FixMe
SQLz = "SELECT Locations.*, lookup_States.StTaxType " & _
       "FROM Locations LEFT JOIN lookup_States ON Locations.LState = lookup_States.StateAbbrev " & _
       "WHERE LocationId =" & Me.LocationList.Column(0) & " " & _
       "ORDER BY Locations.LocationType, Locations.LState, Locations.LCity;"
Me.Form.RecordSource = SQLz
SQLz = "Select AssetID, MLNO, Schedule, [Unit#], Desc From UnitsParkedLocations Where GarageLocationID=" & Me.LocationList.Column(0) & " Order by MLOrig, AssetID;"
Me.UnitsList.RowSource = SQLz
FixMe:
End Sub

Private Sub LState_Click()
    If Nz(Me.LState, "XX") = "XX" Then
        Me.LCity.RowSource = ""
        Me.LCounty.RowSource = ""
        Exit Sub
    End If
    Me.LStateID = Me.LState.Column(3)
    Me.LTaxOption = Me.LState.Column(1)
    Me.LocationTax.Requery
    Dim StateReply As String
    Dim NumRecords As Integer
    StateReply = Nz(LookUpByState(Me.LState), "0|")
    NumRecords = CInt(SplitString(StateReply, "|", 0))
    Me.LCounty = CheckCounty(Me.LCounty, Me.LState, Me.LCity)
    Me.LCity = CheckCity(Me.LCounty, Me.LState, Me.LCity)
    If NumRecords = 0 Then
        Me.LCity.RowSource = ""
        Me.LCounty.RowSource = ""
        Exit Sub
    End If
    Me.LCity.RowSource = SplitString(StateReply, "|", 4)
    Me.LCounty.RowSource = SplitString(StateReply, "|", 5)
End Sub
Private Sub LCity_Click()
    Dim CityReply As String
    Dim NumRecords As Integer
    CityReply = LookUpByCity(Me.LCity, Me.LState)
    If CityReply = "" Then NumRecords = 0 Else NumRecords = CInt(SplitString(CityReply, "|", 0))
    Me.LCounty = CheckCounty(Me.LCounty, Me.LState, Me.LCity)
    If NumRecords = 0 Then
        MsgBox "You've entered a City that is not in the ATLAAS Core, Would you like to add it?", vbCritical, AtlaasVersion
        Exit Sub
    End If
    Me.LCity.RowSource = SplitString(CityReply, "|", 4)
    Me.LCounty.RowSource = SplitString(CityReply, "|", 5)
End Sub
Private Sub LCounty_Click()
Dim CountyReply As String
    Dim NumRecords As Integer
    CountyReply = LookUpByCounty(Nz(Me.LCounty, ""), Me.LState)
    If CountyReply = "" Then NumRecords = 0 Else NumRecords = CInt(SplitString(CountyReply, "|", 0))
    If NumRecords = 0 Then
    MsgBox "You've entered a County that is not in the ATLAAS Core, Would you like to add it?", vbCritical, AtlaasVersion
    Exit Sub
    End If
    Me.LCity.RowSource = SplitString(CountyReply, "|", 4)
    Me.LCounty.RowSource = SplitString(CountyReply, "|", 5)
    If Nz(Me.LCity, "tbd") = "tbd" Or Me.LCity = "" Then Me.LCity = SplitString(CountyReply, "|", 1)
    Me.LCountyID = Me.LCounty.Column(1)
    If CheckCity(Me.LCounty, Me.LState, Me.LCity) <> Me.LCity Then
    MsgBox "The city '" & Me.LCity & "'" & vbNewLine & "is not normaly listed to be within " & Me.LCounty & " county" & vbNewLine & "(the city was not modified, this is just a notice)", vbCritical, AtlaasVersion
    Me.LCity.SetFocus
    End If
    'If NumRecords > 1 Then
    '    MsgBox "Please make sure to select the City and Zip", , AtlaasVersion
    'End If
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
'    Me.LCity = ""
'    Me.LState = ""
 '   Me.LCounty = ""
    Me.LCity.RowSource = SplitString(ZipReply, "|", 4)
    Me.LCounty.RowSource = SplitString(ZipReply, "|", 5)
    If NumRecords > 1 Then
        MsgBox " The Zip you entered has more than one record." & vbNewLine & "Please make sure to select the proper City and County"
    End If
    If Nz(Me.LCity, "") = "" Then Me.LCity = SplitString(ZipReply, "|", 1)
    If Nz(Me.LState, "") = "" Then Me.LState = SplitString(ZipReply, "|", 2)
    If Nz(Me.LCounty, "") = "" Then Me.LCounty = SplitString(ZipReply, "|", 3)
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
'https://www.zip-codes.com/search.asp?srch-type=fuzzy&q=" & [CityName] & "%2C+" & [StateName] & ""

'Private Sub LookUpZipWeb(Zipx As String)
'If Len(Zipx) < 4 Then Exit Sub
'Me.WebBrowser0.Object.Silent = True
'Me.WebBrowser0.Object.Navigate "https://www.zip-codes.com/zip-code/" & Left(Zipx, 5) & "/zip-code-" & Left(Zipx, 5) & ".asp"
'       Call Pause(1)
'    Do While WebBrowser0.Object.Busy
'      DoEvents
'    Loop
''Dim RawString As String
'Dim WebString As String
'Dim CityX, StateX, CountyX, AlaisX, FIPSx As String
'WebString = Me.WebBrowser0.Object.Document.DocumentElement.outerHTML
'WebString = SplitString(SplitString(WebString, "ZIP Code " & Left(Me.LZip, 5) & " Data</h2>", 1), "</table>", 0)
'CityX = SplitString(SplitString(SplitString(WebString, "class=Tips2>City:", 1), "</A>", 0), ".asp" & Chr(34) & ">", 1)
'StateX = SplitString(SplitString(SplitString(WebString, "/state/", 1), " [", 0), ".asp" & Chr(34) & ">", 1)
'CountyX = ProperCase(SplitString(SplitString(SplitString(WebString, "/county/", 1), ",", 0), ".asp" & Chr(34) & ">", 1), 1)
'AlaisX = SplitString(SplitString(SplitString(WebString, "City Alias(es):", 1), "</TD></TR>", 0), "<TD class=info>", 1)
'FIPSx = SplitString(SplitString(WebString, "FIPS Code: ", 1), "" & Chr(34) & " href=", 0)
'Me.LState = StateX
'Me.LCity = CityX
'Me.LCounty = CountyX
'Me.Form.Refresh
'MsgBox Nz(DLookup("TaxCountyID", "vw_TaxKeys", "State='" & StateX & "' and County='" & CountyX & "' "), 0)
'MsgBox Nz(DLookup("TaxStateID", "vw_TaxKeys", "State='" & StateX & "' and County='" & CountyX & "' "), 0)
'MsgBox Nz(DLookup("TaxCityID", "vw_TaxKeys", "State='" & StateX & "' and City='" & CityX & "' "), 0)
''MsgBox StateX & " " & CityX & " " & CountyX & " " & AlaisX & " " & FipsX, vbInformation
'End Sub

Private Sub SaveMe_Click()
Me.Form.Refresh
End Sub

Private Sub SearchAssetId_Change()
    'Me.SearchAssetId = ""
    Me.SearchUnit = ""
    Me.SearchLease = ""
    Me.SearchSchedule = ""
    SQLz = "Select AssetID, MLNO, Schedule, [Unit#], Desc From UnitsParkedLocations Where GarageLocationID=" & Me.LocationList.Column(0) & " " & _
    "AND AssetID Like '*" & Me.SearchAssetID.Text & "*' Order by AssetID;"
    Me.UnitsList.RowSource = SQLz
End Sub
Private Sub SearchLease_Change()
    Me.SearchAssetID = ""
    Me.SearchUnit = ""
    'Me.SearchLease = ""
    Me.SearchSchedule = ""
    SQLz = "Select AssetID, MLNO, Schedule, [Unit#], Desc From UnitsParkedLocations Where GarageLocationID=" & Me.LocationList.Column(0) & " " & _
    "AND MLNo Like '*" & Me.SearchLease.Text & "*' Order by AssetID;"
    Me.UnitsList.RowSource = SQLz
End Sub
Private Sub SearchSchedule_Change()
    Me.SearchAssetID = ""
    Me.SearchUnit = ""
    Me.SearchLease = ""
    'Me.SearchSchedule = ""
    SQLz = "Select AssetID, MLNO, Schedule, [Unit#], Desc From UnitsParkedLocations Where GarageLocationID=" & Me.LocationList.Column(0) & " " & _
    "AND Schedule Like '*" & Me.SearchSchedule.Text & "*' Order by AssetID;"
    Me.UnitsList.RowSource = SQLz
End Sub
Private Sub SearchUnit_Change()
    Me.SearchAssetID = ""
    'Me.SearchUnit = ""
    Me.SearchLease = ""
    Me.SearchSchedule = ""
    SQLz = "Select AssetID, MLNO, Schedule, [Unit#], Desc From UnitsParkedLocations Where GarageLocationID=" & Me.LocationList.Column(0) & " " & _
    "AND [Unit#] Like '*" & Me.SearchUnit.Text & "*' Order by [Unit#];"
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
