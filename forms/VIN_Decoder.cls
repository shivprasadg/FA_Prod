Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz, WebString As String


Private Sub BlankWithUnits_Click()
    SQLz = "Select * From GeoCodeLocations " & _
           "WHERE Units > 0 and (isnull(Longitude ) or Longitude = 0) " & _
           "ORDER BY Client,LocationSn, Units, Units DESC;"
    Me.LocationsList.RowSource = SQLz
End Sub

Private Sub DecodeVIN_Click()
'Me.WebBrowser0.Object.Document.getElementsByName("vin")(0).Value = Me.VINToDecode
    Me.WebBrowser0.Object.Silent = True
    Me.WebBrowser0.Object.Navigate "http://vinfreecheck.com/decode_vin/""" & Me.VinToDecode & """"
    Do While Me.WebBrowser0.Object.ReadyState <> 4: DoEvents: Loop
    '      Application.Wait (Now + TimeValue("0:00:02"))

    WebString = Me.WebBrowser0.Object.Document.body.innerText
    If WebString Like "*Vin too short*" Then
        MsgBox VinToDecode & vbNewLine & "VIN Too Short"
        Exit Sub
    End If
    Dim VinX, YearX, MakeX, ModelX, TrimLevelX, EngineX, CabTypeX, MFGx, SerialX, CheckX, AccuracyX
    'WebString = Split(WebString, "Get Certified VIN Report")(0)
    'WebString = Split(WebString, "Decoded VIN Data")(1)
    VinX = Split(WebString, "VIN Data")(1)
    VinX = Trim(Split(VinX, "YEAR")(0))
    SerialX = Split(WebString, "SEQUENTIAL NUMBER")(1)
    SerialX = Split(SerialX, "VIN CHECK SUM")(0)
    CheckX = Split(WebString, "VIN CHECK SUM")(1)
    CheckX = Split(CheckX, "DATA ACCURACY")(0)
    AccuracyX = Split(WebString, "ACCURACY")(1)
    AccuracyX = Split(AccuracyX, "Vehicle History")(0)
    YearX = Split(WebString, "YEAR")(1)
    YearX = Trim(Split(YearX, "MAKE")(0))
    MakeX = Split(WebString, "MAKE")(1)
    MakeX = Trim(Split(MakeX, "MODEL")(0))
    ModelX = Split(WebString, "MODEL")(1)
    ModelX = Trim(Split(ModelX, "TRIM LEVEL")(0))
    TrimLevelX = Split(WebString, "TRIM LEVEL")(1)
    TrimLevelX = Split(TrimLevelX, "ENGINE")(0)
    EngineX = Split(WebString, "ENGINE")(1)
    EngineX = Split(EngineX, "STYLE")(0)
    CabTypeX = Split(WebString, "STYLE")(1)
    CabTypeX = Split(CabTypeX, "MADE IN")(0)
    MFGx = Split(WebString, "MADE IN")(1)
    MFGx = Split(MFGx, "STEERING")(0)



    MsgBox VinX & vbNewLine & SerialX & vbNewLine & CheckX & vbNewLine & AccuracyX & vbNewLine & YearX & vbNewLine & MakeX & vbNewLine & ModelX & vbNewLine & TrimLevelX & vbNewLine & EngineX & vbNewLine & CabTypeX & vbNewLine & MFGx

End Sub

Private Sub DecodeVIN2_Click()
    Call DownloadVINCSV("https://vpic.nhtsa.dot.gov/decoder/Decoder/ExportToExcel?ModelYear=2014&VIN=" & Me.VinToDecode & "", "C:\FAAS\scripts\VIN.csv")
    Dim FileNum As Integer
    Dim DataLine As String
    WebString = ""
    FileNum = FreeFile()
    Open "C:\FAAS\scripts\VIN.csv" For Input As #FileNum

    While Not EOF(FileNum)
        Line Input #FileNum, DataLine       ' read in CSV file 1 line at a time
        WebString = WebString & DataLine    ' roll each line into a string
    Wend
    Dim VinX, SerialX, CheckX, EngineX, EngMfgX, FuelX, Typex, GvwX, CabTypeX, MakeX, MFGx, YearX, ModelX, DriveX, BrakeX, AxleConfigX, AssyplantX, Notex, HorsePwrX, StateTemX, CountryX

    VinX = Me.VinToDecode
    SerialX = Right(Me.VinToDecode, 8)

    CheckX = Split(WebString, ",")(6)
    CheckX = Split(CheckX, "Possible Values")(0)

    EngineX = Split(WebString, ",")(38)
    EngineX = "L" & EngineX & Split(WebString, ",")(44)
    EngineX = EngineX & Split(WebString, ",")(56) & "L ("
    If Split(WebString, ",")(42) <> "Engine" Then
        EngineX = EngineX & Split(WebString, ",")(42) & " CID)"
    Else
        EngineX = EngineX & ")"
    End If
    EngineX = EngineX & Split(WebString, ",")(48)
    EngineX = Replace(EngineX, "Engine", ", ", 1, 1)
    EngineX = Replace(EngineX, "Engine", "")

    FuelX = Split(WebString, ",")(52)
    FuelX = Split(FuelX, "Engine")(0)

    HorsePwrX = Split(WebString, ",")(62)
    HorsePwrX = Split(HorsePwrX, "Engine")(0)

    EngMfgX = Split(WebString, ",")(76)
    EngMfgX = Split(EngMfgX, "Exterior")(0)

    Typex = Split(WebString, ",")(78)
    Typex = Split(Typex, "Exterior")(0)

    GvwX = Split(WebString, ",")(96)
    GvwX = Split(GvwX, "Exterior")(0)

    CabTypeX = Split(WebString, ",")(124)
    CabTypeX = Split(CabTypeX, "Exterior")(0)

    MakeX = Split(WebString, ",")(134)
    MakeX = Split(MakeX, "General")(0)

    ModelX = Split(WebString, ",")(138)
    ModelX = Split(ModelX, "General")(0)

    YearX = Split(WebString, ",")(140)
    YearX = Split(YearX, "General")(0)

    MFGx = Split(WebString, ",")(136)
    MFGx = Split(MFGx, "General")(0)

    AssyplantX = Split(WebString, ",")(142)
    ';AssyPlantX = AssyPlantX & Split(WebString, ",")(154)

    StateTempX = Split(WebString, ",")(154)
    StateTempX = Split(StateTempX, "General")(0)
    If Len(StateTempX) > 2 Then
        If Not IsNull(DLookup("StateAbbrev", "LookUp_States", "StateLong='" & StateTempX & "'")) Then StateTempX = DLookup("StateAbbrev", "LookUp_States", "StateLong='" & StateTempX & "'")
    End If
    AssyplantX = AssyplantX & ", " & StateTempX & " "

    CountryX = AssyplantX & Split(WebString, ",")(150)
    CountryX = Replace(CountryX, "General", "")

    AssyplantX = AssyplantX & CountryX

    Notex = Split(WebString, ",")(160)
    Notex = Split(Notex, "General")(0)

    BrakeX = Split(WebString, ",")(204)
    BrakeX = Split(BrakeX, "Mechanical")(0)

    DriveX = Split(WebString, ",")(208)
    DriveX = Split(DriveX, "Mechanical")(0)

    AxleConfigX = Split(WebString, ",")(212)
    AxleConfigX = Split(AxleConfigX, "Mechanical")(0)
    Dim SQLz   As String
    SQLz = "INSERT INTO [dbo].[UnitVinDetails] ("
    SQLz = SQLz & "[vin]"
    SQLz = SQLz & ",[unitid]"
    SQLz = SQLz & ",[modelyear]"
    SQLz = SQLz & ",[make]"
    SQLz = SQLz & ",[model]"
    SQLz = SQLz & ",[enginetype]"
    'SQLz = SQLz & ",[transmission]"
    SQLz = SQLz & ",[drivelinetype]"
    SQLz = SQLz & ",[bodytype]"
    SQLz = SQLz & ",[vehicletype]"
    SQLz = SQLz & ",[vehicleclass]"
    SQLz = SQLz & ",[assyplant]"
    SQLz = SQLz & ",[cabtype]"
    ' SQLz = SQLz & ",[trailertype]"
    SQLz = SQLz & ",[fueltype]"
    SQLz = SQLz & ",[country]"
    SQLz = SQLz & ",[horsepower]"
    ' SQLz = SQLz & ",[engineseriescode]"
    SQLz = SQLz & ",[gvwrclass]"
    SQLz = SQLz & ",[axleconfiguration]"
    SQLz = SQLz & ",[manufacturer]"
    SQLz = SQLz & ",[enginemanufacturer]"
    'sqlz = sqlz & ",[plantcode]"
    ' SQLz = SQLz & ",[series]"
    SQLz = SQLz & ",[brakesystem]"
    ' SQLz = SQLz & ",[enginecode]"
    SQLz = SQLz & ",[status]"
    SQLz = SQLz & ",[description])"
    SQLz = SQLz & "VALUES("
    SQLz = SQLz & "'" & VinX & "'"
    SQLz = SQLz & "," & Nz(UnitIDx, 0) & ""
    SQLz = SQLz & "," & CInt(YearX) & ""
    SQLz = SQLz & ",'" & MakeX & "'"
    SQLz = SQLz & ",'" & ModelX & "'"
    SQLz = SQLz & ",'" & EngineX & "'"
    'SQLz = SQLz & ",'" & transmissionx & "'"
    SQLz = SQLz & ",'" & DriveX & "'"
    SQLz = SQLz & ",'" & bodyx & "'"
    SQLz = SQLz & ",'" & Typex & "'"
    SQLz = SQLz & ",'" & Typex & "'"
    SQLz = SQLz & ",'" & AssyplantX & "'"
    SQLz = SQLz & ",'" & CabTypeX & "'"
    '  sqlz = sqlz & ",'" & TrailerX & "'"
    SQLz = SQLz & ",'" & FuelX & "'"
    SQLz = SQLz & ",'" & CountryX & "'"
    SQLz = SQLz & ",'" & HorsePwrX & "'"
    'SQLz = SQLz & ",'" & EngSerX & "'"
    SQLz = SQLz & ",'" & GvwX & "'"
    SQLz = SQLz & ",'" & AxleConfigX & "'"
    SQLz = SQLz & ",'" & MFGx & "'"
    SQLz = SQLz & ",'" & EngMfgX & "'"
    ' SQLz = SQLz & ",'" & PlantCode & "'"
    ' SQLz = SQLz & ",'" & series & "'"
    SQLz = SQLz & ",'" & BrakeX & "'"
    'SQLz = SQLz & ",'" & enginecode & "'"
    SQLz = SQLz & ", " & CInt(Left(CheckX, 1)) & ""
    SQLz = SQLz & ",'" & CheckX & "'"
    SQLz = SQLz & ")"

    'Call GetVinData
    Exit Sub


End Sub
Sub GetVinData()
' open IE, navigate to Decoder Page and loop until fully loaded
    Set ie = CreateObject("InternetExplorer.Application")
    my_url = "https://vpic.nhtsa.dot.gov/decoder/Decoder"
    '/ExportToExcel?" & Me.VINToDecode
    With ie
        .Silent = True
        .Visible = True
        .Navigate my_url
        .Top = 50
        .Left = 530
        .Height = 820
        .Width = 620

        Do Until Not ie.Busy And ie.ReadyState = 4
            DoEvents
        Loop
    End With

    ' Input the VIN
    ie.Document.getElementById("VIN").Value = Me.VinToDecode
    ie.Document.getElementById("ModelYear").Value = "2017"

    ' Click the "Decode" button
    ie.Document.getElementById("btnSubmit").Click

    Do Until Not ie.Busy And ie.ReadyState = 4
        DoEvents
    Loop

    ie.Document.getElementById("btnFullReport").Click
    Do Until Not ie.Busy And ie.ReadyState = 4
        DoEvents
    Loop

   ' https://vpic.nhtsa.dot.gov/decoder/Decoder/ExportToExcel?ModelYear=2017&VIN=3AKGGBDV2HSJB8548
    Dim els    As String
    WebString = ie.Document.body.innerHTML
    ie.Document.querySelector("html").innerHTML
    'WebString = ie.Document.getElementById("tblVIN").Value
    Dim VinX, YearX, MakeX, ModelX, bodyx, TrimLevelX, EngineX, CabTypeX, MFGx, SerialX, CheckX, AccuracyX, WeightX

    VinX = Me.VinToDecode
    SerialX = Right(Me.VinToDecode, 8)

    CheckX = Split(WebString, "VIN CHECK SUM")(1)
    CheckX = Split(CheckX, "DATA ACCURACY")(0)

    AccuracyX = Split(WebString, "Error Code:")(1)
    AccuracyX = Split(AccuracyX, "Manufacturer:")(0)

    MFGx = Split(WebString, "Manufacturer:")(1)
    MFGx = Split(MFGx, "Vehicle Type:")(0)

    YearX = Split(WebString, "Model Year:")(1)
    YearX = Trim(Split(YearX, "Make:")(0))

    MakeX = Split(WebString, "Make:")(1)
    MakeX = Trim(Split(MakeX, "Model:")(0))

    ModelX = Split(WebString, "Model:")(1)
    ModelX = Trim(Split(ModelX, "Body Class: ")(0))

    bodyx = Split(WebString, "Body Class:")(1)
    bodyx = Trim(Split(ModelX, "Other Information")(0))

    TrimLevelX = Split(WebString, "TRIM LEVEL")(1)
    TrimLevelX = Split(TrimLevelX, "ENGINE")(0)

    EngineX = Split(WebString, "ENGINE")(1)
    EngineX = Split(EngineX, "STYLE")(0)
    CabTypeX = Split(WebString, "STYLE")(1)
    CabTypeX = Split(CabTypeX, "MADE IN")(0)

End Sub
Private Sub SearchLocation_Change()
    SQLz = "Select * From GeoCodeLocations " & _
           "WHERE Locationsn Like '*" & Me.SearchLocation.Text & "*' " & _
           "ORDER BY LocationSn, Client, Units DESC;"
    Me.LocationsList.RowSource = SQLz
End Sub

Private Sub SortBlank_Click()
    SQLz = "Select * From GeoCodeLocations ORDER BY Longitude, Client, Units DESC;"
    Me.LocationsList.RowSource = SQLz
End Sub

Private Sub Form_Load()
    Call Locations_Click
End Sub
Private Sub Buyers_Click()
    SQLz = "SELECT ClientID, ClientShNm, IIf(Len([clientAddress1])>4,[clientAddress1] & ' ' & [clientCity] & '  ' & [clientState] & ' ',Null) & IIf(Len([clientZip])>5,Left([clientZip],5) & '-' & Right([clientZip],4),[clientZip]) AS Addr, clientAddress1, clientAddress2, clientCity, clientState, clientZip, Latitude, Longitude " & _
           "FROM Clients WHERE (((IIf(Len([clientAddress1])>4,[clientAddress1] & ' ' & [clientCity] & ' ' & [clientState] & ' ',Null) & IIf(Len([clientZip])>5,Left([clientZip],5) & '-' & Right([clientZip],4),[clientZip])) Is Not Null));"
    Me.ID.ControlSource = "=clientID"
    Me.Form.RecordSource = SQLz

End Sub

Private Sub Locations_Click()
    SQLz = "SELECT LocationID, LocationType, ClientREF, LocationName, Locationsn, IIf(Len([LAddrLine1])>4,[LAddrLine1] & ' ' & [LCity] & '  ' & [LState] & ' ',Null) & IIf(Len([LZip])>5,Left([LZip],5) & '-' & Right([LZip],4),[LZip]) AS Addr, Locations.LAddrLine1, Locations.LAddrLine2, Locations.LAddrLine3, LCity, LState, LZip, Longitude, Latitude FROM Locations;"
    Me.ID.ControlSource = "=LocationID"
    Me.Form.RecordSource = SQLz
End Sub

Private Sub LocationsList_Click()
    On Error GoTo FixMe
    SQLz = "SELECT LocationID, LocationType, ClientREF, LocationName, " & _
           "Locationsn, IIf(Len([LAddrLine1])>4,[LAddrLine1] & ' ' & [LCity] & '  ' & [LState] & ' ',Null) & IIf(Len([LZip])>5,Left([LZip],5) & '-' & Right([LZip],4),[LZip]) AS Addr, " & _
           "LAddrLine1, LAddrLine2, LAddrLine3, LCity, LState, LZip, Longitude, Latitude FROM Locations " & _
           "WHERE LocationID =" & Me.LocationsList.Column(0) & ";"
    Me.Form.RecordSource = SQLz
FixMe:
End Sub

Private Sub LocationsList_DblClick(Cancel As Integer)
    DoCmd.close acForm, "LocationsManager"
    DoCmd.OpenForm "LocationsManager", , , "LocationID = " & Me.LocationsList.Column(0) & "", , , Me.LocationsList.Column(9)
End Sub

Private Sub Map2_Click()
    Me.WebBrowser0.Object.Navigate "http://maps.live.com/default.aspx?where1=33021"
End Sub



Private Sub MapBase_Click()
'Me.WebBrowser0.ControlSource = "=""C:\FAAS\Images\Map.html"""
    Me.WebBrowser0.Object.Navigate "C:\FAAS\Images\Map.html"
    Me.WebBrowser0.Requery
End Sub
Private Sub GeoCoder_Click()
    Dim ix     As Integer
    'Do Until IsNull(Me.LocationID)
    Do While ix < 71
        'Me.WebBrowser0.Object.Navigate "C:\FAAS\Images\Map.html"
        If Me.AddressToSearch & "test" <> "test" Then
            'And Not Me.AddressToSearch Like "*P.O*" Then
            Call MapBase_Click
            Call Pause(1)
            Call SetAddr_Click
            Call Pause(0.5)
            Call GetLatLog_Click
        End If
        DoCmd.GotoRecord Record:=acNext
        ix = ix + 1
    Loop
    'Loop
    'Dim WebX As String
    'WebX = GetBrowserData("GetJavascriptVar")
End Sub

Private Sub MoveNext_Click()
    DoCmd.GotoRecord Record:=acNext
End Sub

Private Sub SetAddr_Click()
    Dim data   As String
    If Not IsNull(Me.AddressToSearch) Then SetBrowserData "ClickGeocode", Me.AddressToSearch
End Sub
Private Sub GetLatLog_Click()
    Me.Latitude = val(WebBrowser0.Object.Document.getElementById("LatX").Value)
    Me.Longitude = val(WebBrowser0.Object.Document.getElementById("LogX").Value)
    Me.LocationsList.Requery
End Sub
Private Function GetBrowserData(ByVal FunctionName As String) As String
' Execute javascript function
    WebBrowser0.Object.Document.parentWindow.execScript (FunctionName & "()")
    ' wait while operation is completed
    While WebBrowser0.Object.Busy
        DoEvents
    Wend
    ' Using DOM, retrieve hidden field value
    GetBrowserData = WebBrowser0.Object.Document.getElementById("LocX").Value

End Function
Private Sub SetBrowserData(ByVal FunctionName As String, ByVal data As String)

' Using DOM, set hidden field value
    WebBrowser0.Object.Document.getElementById("VIN").Value = data

    'Execute javascript function

    WebBrowser0.Object.Document.parentWindow.execScript (FunctionName & "()")

    ' wait while operation is completed
    While WebBrowser0.Object.Busy
        DoEvents
    Wend
End Sub

Private Sub Route1_Click()
    Me.WebBrowser0.ControlSource = "=""http://dev.virtualearth.net/REST/V1/Routes?" & _
                                   "wp.0=San%20Francisco&" & _
                                   "wp.1=San%20Diego&" & _
                                   "du = Mile&" & _
                                   "key=Aq9vHOqzuenBKa_Ll8RrL9RIyTNRmBcnq9iy4a5AYUGK7_AnYrzibzjaFpPW76I8"""
    Me.WebBrowser0.Requery
End Sub
