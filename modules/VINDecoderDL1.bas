Attribute VB_Name = "VINDecoderDL1"
Option Compare Database
Option Explicit
#If VBA7 Then ' Win API Declarations for 32 and 64 bit versions of Office 2010 and later
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
                                           ByVal pCaller As Long, ByVal szURL As String, _
                                           ByVal szFileName As String, _
                                           ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#Else
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
                                           ByVal pCaller As Long, ByVal szURL As String, _
                                           ByVal szFileName As String, _
                                           ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#End If

Function OpenGetVINDialog()
'    If Not VBA.Environ("HOMESHARE") Like "*FA-File*" Then
'        MsgBox "Sorry, You're not on the Fleet Advantage Network", , "Not Local"
'        Exit Function
'    End If
    Dim f      As Object
    Set f = Application.FileDialog(3)
    f.InitialFileName = Environ("Userprofile") & "\Documents"
    f.AllowMultiSelect = False
    f.Filters.Clear
    f.Filters.Add "Excel", "*.xlsx"
    f.Filters.Add "Excel 97", "*.xlx"
    f.Filters.Add "CSV", "*.csv"
    f.Filters.Add "Excel Macro-Enabled", "*.xlsm"
    f.Show
    On Error GoTo FixMe
    ' MsgBox "file choosen = " & f.SelectedItems(1)
    Call DecodeVinList(f.SelectedItems(1), Date)
    Exit Function
FixMe:
End Function
Public Function DecodeVinList(sFile As String, Optional DataDate As String, Optional UpdateEng As Boolean, Optional DecodeFromTemp As Boolean, Optional SaveToServer As Boolean)
    DoCmd.SetWarnings False
    Dim xlApp  As Object
    Dim xlSheet As Object
    Dim xlWB   As Object
    Dim SelectedRange, WorkingRange As Range
    Dim i, UnitYearX, xCount As Integer
    Dim SQLz, DataX(), VinX, XLS, Headerx As String
    Dim lastRow As Long
    Dim TotalRows As Long
    Dim RST    As DAO.Recordset
    Call MessageUser("Open", "Opening Excel", "...this should be quick...")


    If DecodeFromTemp = True Then GoTo Decode ' The temptable alreadly has the Vins to decode.

    '  On Error GoTo Fixme
    Set xlApp = CreateObject("Excel.Application")

    If sFile Like "*.xlsx" Or sFile Like "*.xls" Then sFile = sFile Else sFile = sFile & ".xlsx"
    Set xlApp = CreateObject("Excel.Application")
    With xlApp
        .Visible = True
        Set xlWB = .Workbooks.Open(sFile, , False)
        Call MessageUser("Update", "Getting VINs", "...the original file will be changed...")

        With .ActiveSheet    ' Remove dups, get data rows from source
            lastRow = .Range("A" & .rows.Count).End(xlUp).Row
            .Range("A2:A" & lastRow).RemoveDuplicates
            lastRow = .Range("A" & .rows.Count).End(xlUp).Row
        End With
        xlApp.Application.Range("A1") = "VIN"
        xlApp.Application.Range("B1") = "UnitYear"
        '    xlApp.Application.DisplayAlerts = False
        'xlApp.ActiveWorkbook.SaveAs filename:=sFile
        xlApp.Application.ActiveWorkbook.Save
        xlApp.Application.DisplayAlerts = True
        xlApp.Application.ActiveWorkbook.close
        xlApp.Application.Quit
        TotalRows = TotalRows + lastRow
        DoCmd.SetWarnings False
        DoCmd.RunSQL ("Delete * From Unit_Vin_temp")
        DoCmd.RunSQL ("Delete * From UnitVINResults Where ClientId=0")
        DoCmd.SetWarnings True
        '   xlApp.Application.Save
        DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, "Unit_Vin_temp", sFile, True, "Sheet1!A1:B" & lastRow & ""
        GoTo Decode
        For i = 2 To lastRow
            .Application.Range("A" & i).Select      'Read VIN
            If Len(.Application.Selection.Text) = 0 Then
                VinX = "1"
            Else
                VinX = Nz(Trim(.Application.Selection.Text), "1")
            End If
            .Application.Range("B" & i).Select      'Read Year
            If Len(.Application.Selection.Text) = 0 Then
                UnitYearX = 0
            Else
                UnitYearX = Nz(Trim(.Application.Selection.Text), 0)
            End If
            SQLz = "INSERT INTO Unit_Vin_temp (ClientID, VIN, UnitYear)"
            SQLz = SQLz & "SELECT 0,'" & VinX & "', " & UnitYearX & " "
            DoCmd.SetWarnings False
            DoCmd.RunSQL (SQLz)
            DoCmd.SetWarnings True
        Next i
    End With
    xlApp.Application.Quit
Decode:
    If SaveToServer = True Then
        SQLz = "SELECT ClientID, UnitID, UnitYear, VIN FROM Unit_VIN_Temp ;"    ' Where ClientID=0;"
    Else
        SQLz = "SELECT ClientID, UnitYear, VIN FROM Unit_VIN_Temp ;"    'Where ClientID=0;"
    End If
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    If RST.RecordCount > 0 Then
        RST.MoveLast
        RST.MoveFirst
        xCount = RST.RecordCount
        i = 1
        Call MessageUser("Update", "Getting VIN Data", "working...")
        For i = 1 To xCount
            Call MessageUser("Update", "VIN: " & RST!VIN, "Unit " & i & " Of " & xCount)
            If SaveToServer = True Then
                Call VinDecoder(RST!UnitID, Nz(RST!VIN, "NoVIN"), RST!UnitYear, True, False, RST!clientID)
                If UpdateEng = True Then
                Dim VDUEng As String
                VDUEng = Nz(DLookup("EngX", "vw_VinDecodedEngineUnit", "UnitID=" & RST!UnitID & " "), "Not Provided")
                    SQLz = "UPDATE Units "
                    'SQLz = SQLz & "INNER JOIN vw_VinDecodedEngineUnit As V ON Units.UnitID = V.UnitID "
                    SQLz = SQLz & "SET Units.UnitEngineType = '" & VDUEng & "' "
                    SQLz = SQLz & "WHERE Units.UnitID=" & RST!UnitID & ";"
                    DoCmd.SetWarnings False
                    DoCmd.RunSQL (SQLz)
                    DoCmd.SetWarnings True
                    Call MessageUser("Update", "Set Eng For VIN: " & RST!VIN, "Unit " & i & " Of " & xCount)
                End If
            Else
                If Nz(RST!UnitYear, 0) > 0 Then
                    Call VinDecoder(0, Nz(RST!VIN, "NoVIN"), RST!UnitYear, True, True, Nz(RST!clientID, 0))
                Else
                    Call VinDecoder(0, Nz(RST!VIN, "NoVIN"), , True, True, Nz(RST!clientID, 0))
                End If
            End If
            RST.MoveNext
            'i = i + 1
        Next i
    Else
        GoTo FixMe
    End If

    Headerx = "Decoded VINS - " & Format(Date, "mm-dd-yy")
    XLS = "|C,TL|D,BC|E,TX|J,TC|"
    SQLz = "Select * From UnitVINResults Where ClientId=0"
    SendReport2Excel SQLz, "VINResults", Headerx, True, XLS, 0, True
    Call MessageUser("Close")
    Exit Function
FixMe:
    xlApp.Application.Quit
    Call MessageUser("Close")
    MsgBox "This file doesn't have the proper VIN data" & vbNewLine & "Erorr Code: " & Err, vbCritical, "ATLAAS Info"
End Function


Sub DeleteFile(ByVal FileToDelete As String)
    If FileExists(FileToDelete) Then
        SetAttr FileToDelete, vbNormal  ' remove readonly attribute, if set
        Kill FileToDelete               ' delete the file
    End If
End Sub
Public Function DecodeVinsbyAssetID(AssetIDs As String)
    Dim ThisAsset As Long
    Dim VinX   As String
    Dim x      As Integer
    Dim Times  As Long
    Call MessageUser("Open")
    Times = UBound(Split(AssetIDs, ",")) + 1  'define start and end of array
    If Times < 1 Then Exit Function
    For x = 1 To Times
        On Error GoTo FixMe
        ThisAsset = CLng(SplitString(AssetIDs, ",", x - 1))
        VinX = DLookup("UnitVin", "Units", "UnitId=" & ThisAsset)
        If VinX <> "" Then
            DoCmd.SetWarnings False
            DoCmd.RunSQL ("Delete From UnitVinDetails Where UnitID=" & ThisAsset & ";")
            DoCmd.SetWarnings True
            Call VinDecoder(ThisAsset, VinX)
            Call MessageUser("Update", "Decoding:" & VinX, "Workin on: " & x & " of " & Times, True)
        Else
            MsgBox "No VIN to Decdode", vbCritical, "No Vin"
        End If
    Next x
    Call SetEngineType(True)
    Call MessageUser("Close")
    Exit Function
FixMe:
    If Err.Number <> 13 Then
        Resume Next
    Else
        Call MessageUser("Close")
        Exit Function
    End If
End Function
Public Function ReadMoveOrder()
    Dim SQLz, MoveResponse As String
    Dim FileNum As Integer
    Dim DataLine As String
    MoveResponse = ""
    FileNum = FreeFile()
    Open "C:\FAAS\scripts\report.csv" For Input As #FileNum

    While Not EOF(FileNum)
      Line Input #FileNum, DataLine ' read in CSV file 1 line at a time
    '    If Not DataLine Like "*Order*" And DataLine <> "" Then
       SQLz = "Insert INTO MoveOrderImport (Order , Serial, Driver, ETAx, RevTerm, origin, Client, Dest, Miles, Commodity, ELD, Towable, InStorage) SELECT "
       SQLz = SQLz & "'" & SplitString(DataLine, ",", 0) & "',"
       SQLz = SQLz & "'" & SplitString(DataLine, ",", 1) & "',"
       SQLz = SQLz & "'" & SplitString(DataLine, ",", 2) & "',"
       SQLz = SQLz & "'" & SplitString(DataLine, ",", 3) & "',"
       SQLz = SQLz & "'" & SplitString(DataLine, ",", 4) & "',"
       SQLz = SQLz & "'" & SplitString(DataLine, ",", 5) & "',"
       SQLz = SQLz & "'" & SplitString(DataLine, ",", 6) & "',"
       SQLz = SQLz & "'" & SplitString(DataLine, ",", 7) & "',"
       SQLz = SQLz & "'" & SplitString(DataLine, ",", 8) & "',"
       SQLz = SQLz & "'" & SplitString(DataLine, ",", 9) & "',"
       SQLz = SQLz & "'" & SplitString(DataLine, ",", 10) & "',"
       SQLz = SQLz & "'" & SplitString(DataLine, ",", 11) & "',"
       SQLz = SQLz & "'" & SplitString(DataLine, ",", 12) & "' "
     '  End If
       MoveResponse = MoveResponse & DataLine    ' roll each line into a string
    Wend
    Close #FileNum
End Function

'  CSV DOWNLOADER — STUB (USE YOUR EXISTING HTTP CODE)
Public Function DownloadVINCSV(URL As String, LocalFilename As String) As Boolean
'    Call DeleteFile(LocalFilename)
    Dim GetCSVFile As Long
    GetCSVFile = URLDownloadToFile(0, URL, LocalFilename, 0, 0)
    If GetCSVFile = 0 Then
        If Dir(LocalFilename) <> vbNullString Then
            DownloadVINCSV = True
        End If
    End If
End Function
Public Function GetNHTSA(AssetID As Long, VINstr As String, Optional ModelYear As Integer, _
    Optional Batch As Boolean, Optional CheckOnly As Boolean, Optional ClientIDx As Integer) As Object

    Dim SQLz As String
    Dim VINResponse As String
    Dim strHeader As String
    'Check If this UnitID AND VIN were decoded before.
    Dim DecodedBefore As Boolean
    If Not IsNull(DLookup("UnitId", "UnitVinDetails", "UnitID = " & AssetID & " AND VIN = '" & VINstr & "'")) Then
        DecodedBefore = True
    Else
        DecodedBefore = False
    End If
    'Delete previous CSV file
    If FileExists("C:\FAAS\scripts\VIN.csv") Then
        SetAttr "C:\FAAS\scripts\VIN.csv", vbNormal
        Kill "C:\FAAS\scripts\VIN.csv"
    End If
    'Build URL
    Dim URL As String
    If ModelYear & 0 <> 0 Then
        URL = "https://vpic.nhtsa.dot.gov/api/vehicles/DecodeVinValues/" & VINstr & _
              "?format=csv&modelyear=" & ModelYear
    Else
        URL = "https://vpic.nhtsa.dot.gov/api/vehicles/DecodeVinValues/" & VINstr & "?format=csv"
    End If
    'Download CSV
    Call DownloadVINCSV(URL, "C:\FAAS\scripts\VIN.csv")
    'Read CSV into string
    Dim FileNum As Integer
    Dim DataLine As String
    VINResponse = ""

    FileNum = FreeFile()
    Open "C:\FAAS\scripts\VIN.csv" For Input As #FileNum
    While Not EOF(FileNum)
        Line Input #FileNum, DataLine
        VINResponse = VINResponse & DataLine & vbCrLf
    Wend
    Close #FileNum
    'Extract header + value row
    strHeader = Split(VINResponse, vbCrLf)(0)
    VINResponse = Split(VINResponse, vbCrLf)(1)
    'Return dictionary
    Set GetNHTSA = BuildVINDictionary(strHeader, VINResponse)

End Function
Public Function RoundUp(Number As Double) As Long
    RoundUp = -Int(-Number)
End Function

Public Function VinDecoder(AssetID As Long, VINstr As String, Optional ModelYear As Integer, _
    Optional Batch As Boolean, Optional CheckOnly As Boolean, Optional ClientIDx As Integer)
    Dim dict As Object
    Set dict = GetNHTSA(AssetID, VINstr, ModelYear, , , ClientIDx)
    Dim SQLz As String
    Dim SeriesX, CabTypeX, EngineX, CheckX As String

    If DCount("UnitID", "UnitVinDetails", "UnitID =" & AssetID) > 0 Then
    SQLz = "Delete From UnitVinDetails Where UnitID =" & AssetID & ";"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
    End If

    SeriesX = Trim(Trim(FixUTF8(dict("series") & " " & dict("series2"))) & " " & FixUTF8(dict("trim")))

    CabTypeX = ""
    If SeriesX Like "*Sleeper*" Then CabTypeX = "Sleeper" & CabTypeX
    If SeriesX Like "*Slp*" Then CabTypeX = "Sleeper" & CabTypeX
    If SeriesX Like "*Day Cab*" Then CabTypeX = "Daycab" & CabTypeX
    If SeriesX Like "*DayCab*" Then CabTypeX = "Daycab" & CabTypeX
    If SeriesX Like "*DC*" Then CabTypeX = "Daycab" & CabTypeX
    If CabTypeX & "test" <> "test" Then CabTypeX = CabTypeX & " " & dict("bodycabtype")

    If dict("engineconfiguration") = "In-Line" Then EngineX = "L" Else EngineX = "V"
    If dict("displacementl") & "test" <> "test" Then
        EngineX = EngineX & Trim(dict("enginecylinders") & ", " & dict("displacementl") & "L (" & Round(dict("displacementci"), 2) & " CID)")
    End If
    If EngineX = "V" Then EngineX = "N/A"

    CheckX = dict("errortext")
    If CInt("0" & Left(CheckX, 2)) > 1 Or CheckX = "" Then
        DoCmd.SetWarnings True
        If IsNull(Batch) Or Batch = 0 Then MsgBox "This Unit's VIN failed to decode:" & vbNewLine & "Error Code: " & CheckX, vbCritical, "VIN Details Not Updated"
        SQLz = "INSERT INTO vindecodeerrors ( VIN, UnitID, ErrorCode, ErrorDate, ClientID ) " & _
            "SELECT '" & VINstr & "', " & AssetID & ", '" & CheckX & "', #" & Now() & "#, " & Nz(DLookup("ClientID", "vw_SixKeys", "UnitID=" & AssetID), 0) & ";"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
        If CheckX = "" Then Exit Function
    End If

    SQLz = "INSERT INTO [UnitVinDetails] ("
    SQLz = SQLz & "UnitID"
    SQLz = SQLz & ",VIN"
    SQLz = SQLz & ",ClientID"
    SQLz = SQLz & ",assyplant"
    SQLz = SQLz & ",axleconfiguration"
    SQLz = SQLz & ",bodytype"
    SQLz = SQLz & ",brakesystem"
    SQLz = SQLz & ",bustype"
    SQLz = SQLz & ",cabtype"
    SQLz = SQLz & ",country"
    SQLz = SQLz & ",drivelinetype"
    SQLz = SQLz & ",emissionscode"
    SQLz = SQLz & ",enginecode"
    SQLz = SQLz & ",enginemanufacturer"
    SQLz = SQLz & ",enginemodifier"
    SQLz = SQLz & ",engineseriescode"
    SQLz = SQLz & ",enginetype"
    SQLz = SQLz & ",fueltype"
    SQLz = SQLz & ",gvwrclass"
    SQLz = SQLz & ",gvwrprodrange"
    SQLz = SQLz & ",horsepower"
    SQLz = SQLz & ",hybridtype"
    SQLz = SQLz & ",make"
    SQLz = SQLz & ",manufacturer"
    SQLz = SQLz & ",model"
    SQLz = SQLz & ",modelyear"
    SQLz = SQLz & ",motorcycletypes"
    SQLz = SQLz & ",plantcode"
    SQLz = SQLz & ",restraintsystem"
    SQLz = SQLz & ",series"
    SQLz = SQLz & ",steering"
    SQLz = SQLz & ",trailertype"
    SQLz = SQLz & ",transmission"
    SQLz = SQLz & ",trimlevel"
    SQLz = SQLz & ",vehicleclass"
    SQLz = SQLz & ",vehicletype"
    SQLz = SQLz & ",status"
    SQLz = SQLz & ",description"
    SQLz = SQLz & ",DateDecoded"

    SQLz = SQLz & ") VALUES("
    SQLz = SQLz & "" & AssetID & ""
    SQLz = SQLz & ",'" & dict("vin") & "'"
    SQLz = SQLz & "," & ClientIDx & ""
    SQLz = SQLz & ",'" & Trim(dict("plantcountry") & " " & dict("plantstate") & " " & dict("plantcity")) & "'"
    SQLz = SQLz & ",'" & Trim(dict("drivetype") & " " & dict("axleconfiguration")) & "'"
    SQLz = SQLz & ",'" & Trim(dict("bodyclass vehicletype")) & "'"
    SQLz = SQLz & ",'" & dict("brakesystemtype") & "'"
    SQLz = SQLz & ",'" & dict("bustype ") & "'"
    SQLz = SQLz & ",'" & CabTypeX & "'"
    SQLz = SQLz & ",'" & dict("plantcountry") & "'"
    SQLz = SQLz & ",'" & dict("drivetype") & "'"
    SQLz = SQLz & ",'" & dict("edr = ") & "'"
    SQLz = SQLz & ",'" & dict("engineconfiguration") & "'"
    SQLz = SQLz & ",'" & dict("enginemanufacturer") & "'"
    SQLz = SQLz & ",'" & dict("engineconfiguration") & "'"
    SQLz = SQLz & ",'" & dict("enginemodel") & "'"
    SQLz = SQLz & ",'" & EngineX & "'"
    SQLz = SQLz & ",'" & Trim(dict("fueltypeprimary") & " " & dict("chargerpowerkw")) & "'"
    SQLz = SQLz & ",'" & dict("gvwr") & "'"
    SQLz = SQLz & ",'" & dict("gvwr_to") & "'"
    SQLz = SQLz & ",'" & Trim(dict("enginehp") & " / " & dict("enginehp_to")) & "'"
    SQLz = SQLz & ",'" & Trim(Trim(dict("evdriveunit") & " " & dict("chargerlevel")) & " " & dict("electrificationlevel")) & "'"
    SQLz = SQLz & ",'" & ProperCase(dict("make"), 1) & "'"
    SQLz = SQLz & ",'" & dict("manufacturer") & "'"
    SQLz = SQLz & ",'" & dict("model") & "'"
    SQLz = SQLz & ",'" & dict("modelyear") & "'"
    SQLz = SQLz & ",'" & dict("motorcyclechassistype") & "'"
    SQLz = SQLz & ",'" & Trim(dict("plantcompanyname") & " " & dict("plantstate")) & "'"
    SQLz = SQLz & ",'" & Trim(dict("driverassist") & " " & dict("forwardcollisionwarning")) & "'"
    SQLz = SQLz & ",'" & SeriesX & "'"                         'series
    SQLz = SQLz & ",'" & dict("steeringlocation") & "'"
    SQLz = SQLz & ",'" & Trim(dict("trailerbodytype trailerlength trailertype")) & "'"
    SQLz = SQLz & ",'" & Trim(dict("transmissionspeeds") & " " & dict("transmissionstyle")) & "'"
    SQLz = SQLz & ",'" & Trim(dict("trim") & " " & dict("trim2")) & "'"
    SQLz = SQLz & ",'" & dict("vehicletype") & "'"
    SQLz = SQLz & ",'" & dict("bodyclass") & "'"
    SQLz = SQLz & ",'" & Trim(SplitString(dict("errortext"), "-", 0)) & "'"
    SQLz = SQLz & ",'" & dict("errortext") & "'"
    SQLz = SQLz & ",'" & Now & "'"
    SQLz = SQLz & ")"
    SQLz = Replace(SQLz, "''", "NULL")
' Insert Record
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    If EngineX <> "N/A" Then Call SetEngineType(True)
   '  Dim k As Variant
   '  For Each k In dict.Keys
   '     Debug.Print k & " = " & dict(k)
   '  Next k

End Function
Public Function FixUTF8(s As String) As String
    s = Replace(s, "â€", "”")   ' right double quote
    s = Replace(s, "â€œ", "“")   ' left double quote
    s = Replace(s, "â€™", "’")   ' apostrophe
    s = Replace(s, "â€“", "–")   ' en dash
    s = Replace(s, "â€”", "—")   ' em dash
    s = Replace(s, "â€¦", "…")   ' ellipsis
    FixUTF8 = s
End Function


Public Function BuildVINDictionary(strHeader As String, strValues As String) As Object
    Dim headers As Variant
    Dim values As Variant
    Dim dict As Object
    Dim i As Long

    headers = ParseCSVLine(strHeader)
    values = ParseCSVLine(strValues)

    Set dict = CreateObject("Scripting.Dictionary")

    For i = 0 To UBound(headers)
        If i <= UBound(values) Then
            dict(headers(i)) = values(i)
        Else
            dict(headers(i)) = Null
        End If
    Next i

    Set BuildVINDictionary = dict
End Function

Public Function ParseCSVLine(ByVal line As String) As Variant
    Dim Result As New Collection
    Dim i As Long, ch As String
    Dim inQuotes As Boolean
    Dim current As String

    For i = 1 To Len(line)
        ch = Mid$(line, i, 1)

        Select Case ch
            Case """"   ' toggle quote state
                inQuotes = Not inQuotes

            Case ","    ' comma
                If inQuotes Then
                    current = current & ch
                Else
                    Result.Add current
                    current = ""
                End If

            Case Else
                current = current & ch
        End Select
    Next i

    Result.Add current

    Dim arr() As String
    ReDim arr(0 To Result.Count - 1)

    For i = 1 To Result.Count
        arr(i - 1) = Result(i)
    Next i

    ParseCSVLine = arr
End Function

Public Function VinDecoder_OLD(AssetID As Long, VINstr As String, Optional ModelYear As Integer, _
    Optional Batch As Boolean, Optional CheckOnly As Boolean, Optional ClientIDx As Integer)
    Dim SQLz, VINResponse As String

    'Check If this UnitID AND VIN where decoded before.
    Dim DecodedBefore As Boolean
    If Not IsNull(DLookup("UnitId", "UnitVinDetails", "UnitID = " & AssetID & " AND VIN = '" & VINstr & "'")) Then DecodedBefore = True Else DecodedBefore = False
    'Delete the previous CSV file
    If FileExists("C:\FAAS\scripts\VIN.csv") Then
        SetAttr "C:\FAAS\scripts\VIN.csv", vbNormal
        ' Then delete the file
        Kill "C:\FAAS\scripts\VIN.csv"
    End If
    'Check If were also sending the Model Year along with VIN decode request
    If ModelYear & 0 <> 0 Then
    '    Call DownloadVINCSV("https://vpic.nhtsa.dot.gov/decoder/Decoder/ExportToExcel?ModelYear=" & ModelYear & "&VIN=" & VINstr & "", "C:\FAAS\scripts\VIN.csv")
        Call DownloadVINCSV("https://vpic.nhtsa.dot.gov/api/vehicles/DecodeVinValues/" & VINstr & "?format=csv&modelyear=" & ModelYear & "", "C:\FAAS\scripts\VIN.csv")
    Else
    '    Call DownloadVINCSV("https://vpic.nhtsa.dot.gov/decoder/Decoder/ExportToExcel?VIN=" & VINstr & "", "C:\FAAS\scripts\VIN.csv")
        Call DownloadVINCSV("https://vpic.nhtsa.dot.gov/api/vehicles/DecodeVinValues/" & VINstr & "?format=csv", "C:\FAAS\scripts\VIN.csv")
    End If

    ' The request response is saved as CSV file and read into a string here
    Dim FileNum As Integer
    Dim DataLine As String
    VINResponse = ""
    FileNum = FreeFile()
    Open "C:\FAAS\scripts\VIN.csv" For Input As #FileNum

    While Not EOF(FileNum)
        Line Input #FileNum, DataLine       ' read in CSV file 1 line at a time
        VINResponse = VINResponse & DataLine & vbCrLf   ' roll each line into a string
    Wend
    Close #FileNum

    Dim SerialX, CheckX, enginetypeX, EngMfgX, EngSeriesX, FuelX, Typex, ChassisX, GvwX, CabTypeX, TrailerTypeX, MakeX, MFGx, YearX, ModelX, DriveX, _
        BrakeX, AxleConfigX, AssyplantX, PlantCodeX, Notex, HorsePwrX, StateTemX, CountryX, StateTempX, bodyx, SeriesX, TransmissionX, RestraintSystemX, Tempx As String
    If ClientIDx + 1 = 1 Then ClientIDx = 0

    SerialX = Right(VINstr, 8)
    CheckX = Split(VINResponse, "Error Code,")(1)
    CheckX = Split(CheckX, "Possible Values")(0)
    If CheckX Like "*;*" Then CheckX = Split(CheckX, ";")(0)
    'Engine Type
    enginetypeX = Split(VINResponse, "Cylinders,")(1)
    enginetypeX = Split(enginetypeX, "Engine")(0)
    If enginetypeX <> "" Then enginetypeX = "L" & Trim(enginetypeX) & ","

    Tempx = Split(VINResponse, "(L),")(1)
    Tempx = Split(Tempx, "Engine")(0)
    If Tempx <> "" Then Tempx = Round(CDbl(Tempx), 2)
    If Tempx <> "" Then enginetypeX = enginetypeX & " " & Tempx & "L"
    Tempx = ""
    Tempx = Split(VINResponse, "(CI),")(1)
    Tempx = Split(Tempx, "Engine")(0)
    If Tempx <> "" Then
        Tempx = Round(CDbl(Tempx), 2)
        enginetypeX = enginetypeX & " (" & Tempx & " CID)"
    End If
    'Engine SeriesX
    EngSeriesX = Split(VINResponse, "Engine Model,")(1)
    EngSeriesX = Split(EngSeriesX, "Engine,")(0)
    'Engine MFG
    EngMfgX = Split(VINResponse, "Engine Manufacturer,")(1)
    EngMfgX = Split(EngMfgX, "Exterior")(0)
    'Fuel
    FuelX = Split(VINResponse, "Fuel Type - Primary,")(1)
    FuelX = Split(FuelX, "Engine,")(0)
    'HorsePower
    On Error Resume Next     'File change 2/1/2021 - Trap error
    HorsePwrX = Split(VINResponse, "Brake (hp),")(1)
    If HorsePwrX = "" Then HorsePwrX = Split(VINResponse, "Engine Brake (hp) From,")(1)

    HorsePwrX = Split(HorsePwrX, "Engine")(0)
    Tempx = Split(VINResponse, "Brake (hp) up to,")(1)
    If Tempx = "" Then Tempx = Split(VINResponse, "Engine Brake (hp) To,")(1)
    Tempx = Split(Tempx, "Engine")(0)
    If Tempx <> "" Then HorsePwrX = HorsePwrX & " / " & Tempx & ""
    'Chassis
    ChassisX = Split(VINResponse, "Body Class,")(1)
    ChassisX = Split(ChassisX, "Exterior")(0)
    Typex = ChassisX
    'TrailerType
    TrailerTypeX = Split(VINResponse, "Trailer Body Type,")(1)
    TrailerTypeX = Split(TrailerTypeX, "Exterior")(0)
    Tempx = Split("Trailer Length (feet),")(1)
    Tempx = Split(Tempx, "Exterior")(0)
    If Tempx <> "" Then TrailerTypeX = TrailerTypeX & " (" & Tempx & ")"
    Tempx = ""
    'GVWeight
    GvwX = Split(VINResponse, "Weight Rating,")(1)
    If GvwX = "" Then GvwX = Split(VINResponse, "Gross Vehicle Weight Rating From,")(1)
    GvwX = Split(GvwX, "Exterior")(0)
    On Error GoTo 0  ' end Trap
    'CabType
    CabTypeX = Split(VINResponse, "Cab Type,")(1)
    CabTypeX = Split(CabTypeX, "Exterior")(0)

    MakeX = Split(VINResponse, "Make,")(1)
    MakeX = Split(MakeX, "General")(0)
    'MakeX = ProperCase(MakeX, 1)

    ModelX = Split(VINResponse, ",Model,")(1)
    ModelX = Split(ModelX, "General")(0)

    YearX = Split(VINResponse, "Model Year,")(1)
    YearX = Split(YearX, "General")(0)
    If YearX = "" Then YearX = "0"
    MFGx = Split(VINResponse, "Manufacturer Name,")(1)
    MFGx = Split(MFGx, "General")(0)

    AssyplantX = Split(VINResponse, "Plant Company Name,")(1)
    AssyplantX = Split(AssyplantX, "General")(0)
    If AssyplantX <> "" Then AssyplantX = AssyplantX & " "
    'AssyPlantX = Replace(AssyPlantX, ",", "", 1)

    AssyplantX = Split(VINResponse, "Plant City,")(1)
    AssyplantX = Trim(Split(AssyplantX, "General")(0))

    StateTempX = Split(VINResponse, "Plant State,")(1)
    StateTempX = Split(StateTempX, "General")(0)

    If Len(StateTempX) > 2 Then
        If Not IsNull(DLookup("StateAbbrev", "LookUp_States", "StateLong='" & StateTempX & "'")) Then StateTempX = DLookup("StateAbbrev", "LookUp_States", "StateLong='" & StateTempX & "'")
    End If
    If StateTempX <> "" And AssyplantX <> "" Then AssyplantX = AssyplantX & ", " & StateTempX & " "
    If StateTempX <> "" And AssyplantX = "" Then AssyplantX = StateTempX
    CountryX = Split(VINResponse, "Plant Country,")(1)
    CountryX = Split(CountryX, "General")(0)
    On Error GoTo FixPlant
    PlantCodeX = Split(VINResponse, "Manufacturer Id,")(1)
    PlantCodeX = Split(PlantCodeX, "General")(0)
    ' API changed fix applied 4/20/2021
FixPlant:
    On Error GoTo 0
    If PlantCodeX = "" Then PlantCodeX = Split(Split(VINResponse, "General,Plant State,")(1), "General,Trim2,")(0)

    SeriesX = Split(VINResponse, "Series2,")(1)
    SeriesX = Split(SeriesX, "General")(0)
    Tempx = Split(VINResponse, "Series,")(1)
    Tempx = Split(Tempx, "General")(0)
    If Tempx <> "" And SeriesX <> "" Then SeriesX = SeriesX & " (" & Tempx & ")"
    If Tempx <> "" And SeriesX = "" Then SeriesX = Tempx
    If SeriesX <> "" Then ModelX = ModelX & " - " & SeriesX & ""
    Tempx = ""
    Tempx = Split(VINResponse, "General,Trim,")(1)
    Tempx = Split(Tempx, "General,")(0)
    Notex = Split(VINResponse, "General,Note,")(1)
    Notex = Split(Notex, "General,")(0)
    Notex = Notex & Tempx
    Tempx = ""

    If VINResponse Like "*Sleeper*" Then CabTypeX = "Sleeper " & CabTypeX
    If VINResponse Like "*Slp*" Then CabTypeX = "Sleeper " & CabTypeX
    If VINResponse Like "*Day Cab*" Then CabTypeX = "Daycab " & CabTypeX
    If VINResponse Like "*DayCab*" Then CabTypeX = "Daycab " & CabTypeX
    If VINResponse Like "*DC*" Then CabTypeX = "Daycab " & CabTypeX

    BrakeX = Split(VINResponse, "Brake System Type,")(1)
    BrakeX = Split(BrakeX, "Mechanical")(0)

    DriveX = Split(VINResponse, "Drive Type,")(1)
    DriveX = Split(DriveX, "Mechanical")(0)

    AxleConfigX = Split(VINResponse, "Axles,")(1)
    AxleConfigX = Split(AxleConfigX, "Mechanical")(0)
    If AxleConfigX <> "" Then AxleConfigX = AxleConfigX & " "

    AxleConfigX = Split(VINResponse, "Axle Configuration,")(1)
    AxleConfigX = Split(AxleConfigX, "Mechanical")(0)

    'Transmission
    TransmissionX = Split(VINResponse, "Transmission Style,")(1)
    TransmissionX = Split(TransmissionX, "Mechanical")(0)
    Tempx = Split(VINResponse, "Transmission SPeeds,")(1)
    Tempx = Split(Tempx, "Passive")(0)

    If Tempx <> "" And TransmissionX <> "" Then TransmissionX = TransmissionX & " (" & Tempx & ")"
    If Tempx <> "" And TransmissionX = "" Then TransmissionX = Tempx
    Tempx = ""

    RestraintSystemX = Split(VINResponse, "Passive Safety System,")(1)
    RestraintSystemX = Replace(RestraintSystemX, "Passive Safety System / Air Bag Location,", "")
    RestraintSystemX = Replace(RestraintSystemX, "Passive Safety System,", "")
    RestraintSystemX = Replace(RestraintSystemX, "Curtain Air Bag Locations,", "")
    RestraintSystemX = Replace(RestraintSystemX, "Seat Cushion Air Bag Locations,", "")
    RestraintSystemX = Replace(RestraintSystemX, "Front Air Bag Locations,", "")
    RestraintSystemX = Replace(RestraintSystemX, "Knee Air Bag Locations,", "")
    RestraintSystemX = Replace(RestraintSystemX, "Side Air Bag Locations,", "")
    RestraintSystemX = Replace(RestraintSystemX, "Pretensioner,", "")
    RestraintSystemX = Replace(RestraintSystemX, "Seat Belts Type,", "")
    RestraintSystemX = Replace(RestraintSystemX, "Other Restraint System Info,", "")

    ' Call MessageUser("Update", "VIN: " & VINstr, "Make:" & ProperCase(MakeX, 1) & " Check:" & CheckX)

    If CInt(Left(CheckX, 2)) > 1 Or CheckX = "" Then
        DoCmd.SetWarnings True
        If IsNull(Batch) Or Batch = 0 Then MsgBox "This Unit's VIN failed to decode:" & vbNewLine & "Error Code: " & CheckX, vbCritical, "VIN Details Not Updated"
        SQLz = "INSERT INTO vindecodeerrors ( VIN, UnitID, ErrorCode, ErrorDate, ClientID ) " & _
            "SELECT '" & VINstr & "', " & AssetID & ", '" & CheckX & "', #" & Now() & "#, " & Nz(DLookup("ClientID", "vw_SixKeys", "UnitID=" & AssetID), 0) & ";"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
        If CheckX = "" Then Exit Function
    End If

    If DecodedBefore = True And CInt(Left(CheckX, 1)) < 2 Then
        SQLz = "Delete From UnitVinDetails Where vin = '" & VINstr & "';"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
    End If
    Notex = Notex & vbNewLine & CheckX
    If CheckOnly = True Then SQLz = "INSERT INTO [UnitVINResults] (" Else SQLz = "INSERT INTO [UnitVinDetails] ("
    SQLz = SQLz & " [ClientID]"
    SQLz = SQLz & ",[vin]"
    SQLz = SQLz & ",[UnitID]"
    SQLz = SQLz & ",[modelyear]"
    SQLz = SQLz & ",[make]"
    SQLz = SQLz & ",[model]"
    SQLz = SQLz & ",[enginetype]"
    SQLz = SQLz & ",[drivelinetype]"
    SQLz = SQLz & ",[bodytype]"
    SQLz = SQLz & ",[vehicletype]"
    SQLz = SQLz & ",[vehicleclass]"
    SQLz = SQLz & ",[assyplant]"
    SQLz = SQLz & ",[cabtype]"
    SQLz = SQLz & ",[fueltype]"
    SQLz = SQLz & ",[country]"
    SQLz = SQLz & ",[horsepower]"
    SQLz = SQLz & ",[gvwrclass]"
    SQLz = SQLz & ",[axleconfiguration]"
    SQLz = SQLz & ",[manufacturer]"
    SQLz = SQLz & ",[enginemanufacturer]"
    SQLz = SQLz & ",[brakesystem]"
    SQLz = SQLz & ",[status]"
    SQLz = SQLz & ",[description]"
    SQLz = SQLz & ",[series]"
    SQLz = SQLz & ",[trailertype]"
    SQLz = SQLz & ",[transmission]"
    SQLz = SQLz & ",[plantcode]"
    SQLz = SQLz & ",[engineseriescode]"
    SQLz = SQLz & ",[RestraintSystem]"
    SQLz = SQLz & ",[DateDecoded]"
    'SQLz = SQLz & ",[wheelbase]"
    'SQLz = SQLz & ",[enginecode]"
    'SQLz = SQLz & ",[bustype]"
    'SQLz = SQLz & ",[carlinecode]"
    'SQLz = SQLz & ",[chassis]"

    SQLz = SQLz & ") VALUES("
    SQLz = SQLz & " " & ClientIDx & ""
    SQLz = SQLz & ",'" & VINstr & "'"
    SQLz = SQLz & ", " & AssetID & ""
    SQLz = SQLz & ", " & CInt(YearX) & ""
    SQLz = SQLz & ",'" & ProperCase(MakeX, 1) & "'"
    SQLz = SQLz & ",'" & ModelX & "'"
    SQLz = SQLz & ",'" & enginetypeX & "'"
    SQLz = SQLz & ",'" & DriveX & "'"
    SQLz = SQLz & ",'" & bodyx & "'"
    SQLz = SQLz & ",'" & Typex & "'"
    SQLz = SQLz & ",'" & Typex & "'"
    SQLz = SQLz & ",'" & AssyplantX & "'"
    SQLz = SQLz & ",'" & CabTypeX & "'"
    SQLz = SQLz & ",'" & FuelX & "'"
    SQLz = SQLz & ",'" & CountryX & "'"
    SQLz = SQLz & ",'" & HorsePwrX & "'"
    SQLz = SQLz & ",'" & GvwX & "'"
    SQLz = SQLz & ",'" & AxleConfigX & "'"
    SQLz = SQLz & ",'" & ProperCase(MFGx, 1) & "'"
    SQLz = SQLz & ",'" & ProperCase(EngMfgX, 1) & "'"
    SQLz = SQLz & ",'" & BrakeX & "'"
    SQLz = SQLz & ", " & CInt(Left(CheckX, 1)) & ""
    SQLz = SQLz & ",'" & Notex & "'"
    SQLz = SQLz & ",'" & SeriesX & "'"
    SQLz = SQLz & ",'" & TrailerTypeX & "'"
    SQLz = SQLz & ",'" & TransmissionX & "'"
    SQLz = SQLz & ",'" & PlantCodeX & "'"
    SQLz = SQLz & ",'" & EngSeriesX & "'"
    SQLz = SQLz & ",'" & RestraintSystemX & "'"
    SQLz = SQLz & ",#" & Now() & "#"
    'SQLz = SQLz & ",'" & wheelbase & "'"
    'SQLz = SQLz & ",'" & enginecode & "'"
    'SQLz = SQLz & ",'" & bustype & "'"
    'SQLz = SQLz & ",'" & carlinecode & "'"
    'SQLz = SQLz & ",'" & ChassisX & "'"
    SQLz = SQLz & ")"
    ' Insert Record
    SQLz = Replace(SQLz, "'S ", "s ")
    SQLz = Replace(SQLz, "'37'6", "'37ft 6")
    SQLz = Replace(SQLz, "'44'", "'44ft")
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    ' Set Units.UnitEngineType to Standard Format
    If enginetypeX <> "" Then Call SetEngineType(True)
    'Call GetVinData
End Function

Public Function Set_EngMfg(Optional noMsg As Boolean)
   ' Dim SQLz As String
   ' DoCmd.SetWarnings False
    'SQLz = "Update [UnitVinDetails] Set enginemanufacturer = 'Mercedes-Benz' where enginemanufacturer = 'Mercedes'"
   ' DoCmd.RunSQL (SQLz)
   ' SQLz = "Update [UnitVinDetails] Set enginemanufacturer = 'International / Navistar' where enginemanufacturer = 'International/navistar'"
   ' DoCmd.RunSQL (SQLz)
   ' SQLz = "Update [UnitVinDetails] Set enginemanufacturer = 'Detroit Diesel' where enginemanufacturer = 'Detroit' "
   ' DoCmd.RunSQL (SQLz)
   ' SQLz = "Update [UnitVinDetails] Set enginemanufacturer = 'Caterpillar' where enginemanufacturer = 'Cat' "
   ' DoCmd.RunSQL (SQLz)
   ' DoCmd.SetWarnings True
   ' If noMsg = False Then MsgBox "VIN Decode and Engine MFG Fix Done", vbOKOnly, "ATLAAS Core"
End Function
