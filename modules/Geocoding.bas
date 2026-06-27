Attribute VB_Name = "Geocoding"
Option Compare Database

Const pi = 3.14159265358979

Function Distance(lat1, lon1, lat2, lon2, Unit)
  Dim theta, dist
  theta = lon1 - lon2
  dist = Sin(deg2rad(lat1)) * Sin(deg2rad(lat2)) + Cos(deg2rad(lat1)) * Cos(deg2rad(lat2)) * Cos(deg2rad(theta))
  dist = acos(dist)
  dist = rad2deg(dist)
  Distance = dist * 60 * 1.1515
  Select Case UCase(Unit)
    Case "K"
      Distance = Distance * 1.609344
    Case "N"
      Distance = Distance * 0.8684
  End Select
End Function


'
'  This function get the arccos function from arctan function
'
Function acos(Rad)
  If Abs(Rad) <> 1 Then
    acos = pi / 2 - Atn(Rad / Sqr(1 - Rad * Rad))
  ElseIf Rad = -1 Then
    acos = pi
  End If
End Function


'
'  This function converts decimal degrees to radians
'
Function deg2rad(Deg)
    deg2rad = CDbl(Deg * pi / 180)
End Function

'
'  This function converts radians to decimal degrees
'
Function rad2deg(Rad)
    rad2deg = CDbl(Rad * 180 / pi)
End Function

Public Function GetCityStateCounty()
Dim objIE As Object

Set objIE = CreateObject("InternetExplorer.Application")

With objIE
    .Navigate "https://www.zip-codes.com/zip-code/33021/zip-code-33021.asp"
     Do While .Busy
         ' Application.Wait Now + TimeValue("0:00:01")
     Loop
    .Visible = True
End With

'HTML document
Dim doc As Object
Dim WebString As String
Set doc = objIE.Document
WebString = doc.DocumentElement.outerHTML
Dim el As Object
Dim myText As String
For Each el In doc.getElementsByClassName("myClass")
    'put paragrah text in cell A1f
'    Cells(1, 1).Value = el.innerText

    'put your paragraph text in a variable string
    myText = el.innerText
Next el
End Function

Public Function LookUpByZip(ZipcodeX As String) As String
Exit Function
Dim CityX, StateX, CountyX, SQLx, SQLy, SQLz As String
 Dim RST    As DAO.Recordset
    SQLx = "Select State, City, County, Zip, Lat, Long, CountyID From vw_CCSZ_Lookup Where Zip ='" & ZipcodeX & "' Order By State, City;"
    Set RST = CurrentDb.OpenRecordset(SQLx, dbOpenSnapshot)
    On Error Resume Next
    RST.MoveLast
    RST.MoveFirst
    On Error GoTo Cleanup
    SQLy = "Select City From vw_CCSZ_Lookup Where Zip ='" & ZipcodeX & "' Group By City Order By City;"
    SQLz = "Select County, CountyID From vw_CCSZ_Lookup Where Zip ='" & ZipcodeX & "' Group By County Order By County;"
    CityX = RST!City
    StateX = RST!State
    CountyX = RST!County
LookUpByZip = RST.RecordCount & "|" & CityX & "|" & StateX & "|" & CountyX & "|" & SQLy & "|" & SQLz & "|" & RST!Lat & "|" & RST![Long] & "|" & RST![CountyID]
Cleanup:
RST.close
Set RST = Nothing
End Function
Public Function LookUpByState(StateX As String) As String
Dim CityX, CountyX, SQLx, SQLy, SQLz As String
 Dim RST    As DAO.Recordset
    SQLx = "Select * From vw_CCSZ_Lookup Where State ='" & StateX & "' Order By State, County, City;"
    Set RST = CurrentDb.OpenRecordset(SQLx, dbOpenSnapshot)
    On Error Resume Next
    RST.MoveLast
    RST.MoveFirst
    On Error GoTo Cleanup
    SQLy = "Select City, CityID, County, CountyID From vw_CCSZ_Lookup Where State ='" & StateX & "' Group By City, CityID, County, CountyID Order By City, County;"
    SQLz = "Select County, CountyID From vw_CCSZ_Lookup Where State ='" & StateX & "' Group By County, CountyID Order By County;"
    CityX = RST!City
    'StateX = Rst!State
    CountyX = RST!County
LookUpByState = RST.RecordCount & "|" & CityX & "|" & StateX & "|" & CountyX & "|" & SQLy & "|" & SQLz & "|" & RST!Lat & "|" & RST![Long] & "|" & RST![CountyID] & "|" & RST![CityID]
Cleanup:
RST.close
Set RST = Nothing
End Function
Public Function LookUpByCity(CityX As String, StateX As String) As String
Dim CountyX, SQLx, SQLy, SQLz As String
 Dim RST    As DAO.Recordset
    SQLx = "Select * From vw_CCSZ_Lookup Where State ='" & StateX & "' AND City= '" & CityX & "' Order By State, County, City;"
    Set RST = CurrentDb.OpenRecordset(SQLx, dbOpenSnapshot)
    On Error Resume Next
    RST.MoveLast
    RST.MoveFirst
    On Error GoTo Cleanup
    SQLy = "Select City, CityID, County, CountyID From vw_CCSZ_Lookup Where State ='" & StateX & "' Group By City, CityID, County, CountyID Order By City, County;"
    SQLz = "Select County, CountyID From vw_CCSZ_Lookup Where State ='" & StateX & "' AND City= '" & CityX & "' Group By County, CountyID Order By County;"
    'CityX = Rst!City
    'StateX = Rst!State
    CountyX = RST!County
LookUpByCity = RST.RecordCount & "|" & CityX & "|" & StateX & "|" & CountyX & "|" & SQLy & "|" & SQLz & "|" & RST!Lat & "|" & RST![Long] & "|" & RST![CountyID] & "|" & RST![CityID]
Cleanup:
RST.close
Set RST = Nothing
End Function
Public Function LookUpByCounty(CountyX As String, StateX As String) As String
Dim CityX, SQLx, SQLy, SQLz As String
 Dim RST    As DAO.Recordset
    SQLx = "Select * From vw_CCSZ_Lookup Where State ='" & StateX & "' AND County= '" & CountyX & "' Order By State, County, City;"
    Set RST = CurrentDb.OpenRecordset(SQLx, dbOpenSnapshot)
    On Error Resume Next
    RST.MoveLast
    RST.MoveFirst
    On Error GoTo Cleanup
    SQLy = "Select City, CityID, County, CountyID From vw_CCSZ_Lookup Where State ='" & StateX & "' Group By City, CityID, County, CountyID Order By City, County;"
    SQLz = "Select County, CountyID From vw_CCSZ_Lookup Where State ='" & StateX & "' Group By County, CountyID Order By County;"
    CityX = RST!City
    'StateX = Rst!State
    'CountyX = Rst!County
LookUpByCounty = RST.RecordCount & "|" & CityX & "|" & StateX & "|" & CountyX & "|" & SQLy & "|" & SQLz & "|" & RST!Lat & "|" & RST![Long] & "|" & RST![CountyID] & "|" & RST![CityID]
Cleanup:
RST.close
Set RST = Nothing
End Function
Public Function CheckCity(CountyX As String, StateX As String, CityX As String) As String
    CheckCity = "tbd"
    Dim SQLx   As String
    Dim RST    As DAO.Recordset
    SQLx = "Select City From vw_CCSZ_Lookup Where State ='" & StateX & "' "
    If CityX = "" Or CityX = "tbd" Then SQLx = SQLx Else SQLx = SQLx & "AND City= '" & CityX & "' "
    If CountyX = "" Or CountyX = "tbd" Then SQLx = SQLx Else SQLx = SQLx & "AND County= '" & CountyX & "' "
    SQLx = SQLx & " Group By City;"
    Set RST = CurrentDb.OpenRecordset(SQLx, dbOpenSnapshot)
    On Error Resume Next
    RST.MoveLast
    RST.MoveFirst
    On Error GoTo Cleanup
    If RST.RecordCount = 1 Then CheckCity = RST!City
    If RST.RecordCount > 1 Then CheckCity = CityX
Cleanup:
    RST.close
    Set RST = Nothing
End Function
Public Function CheckCounty(CountyX As String, StateX As String, CityX As String) As String
    CheckCounty = "tbd"
    Dim SQLx   As String
    Dim RST    As DAO.Recordset
    SQLx = "Select County From vw_CCSZ_Lookup Where State ='" & StateX & "' "
    If CityX = "" Or CityX = "tbd" Then SQLx = SQLx Else SQLx = SQLx & "AND City= '" & CityX & "' "
    If CountyX = "" Or CountyX = "tbd" Then SQLx = SQLx Else SQLx = SQLx & "AND County= '" & CountyX & "' "
    SQLx = SQLx & " Group By County;"
    Set RST = CurrentDb.OpenRecordset(SQLx, dbOpenSnapshot)
    On Error Resume Next
    RST.MoveLast
    RST.MoveFirst
    On Error GoTo Cleanup
    If RST.RecordCount = 1 Then CheckCounty = RST!County
    If RST.RecordCount > 1 Then CheckCounty = CountyX
Cleanup:
    RST.close
    Set RST = Nothing
End Function
Public Sub GetWebtexts(CityName As String, StateName As String)
   ' Dim pRS As New ADODB.Recordset
    Dim basepath As String
    Dim site_no As String
    Dim inurl As String
    Dim outpath As String

    basepath = "C:\Faas\" & FAUserName & "\Temp\"
   ' pRS.Open "gagelist", CurrentProject.Connection, adOpenStatic, , adCmdTable
   ' pRS.MoveFirst
   ' Do Until pRS.EOF = True
      '  site_no = pRS("site_no").Value
        inurl = "https://www.zip-codes.com/search.asp?srch-type=fuzzy&q=" & CityName & "%2C+" & StateName & ""
      '  inurl = "http://waterdata.usgs.gov/tx/nwis/dv?cb_00060=on&format=rdb&begin_date=2008-10-15&end_date=2009-10-15&referred_module=sw&site_no=" & site_no
        outpath = "C:\Faas\" & FAUserName & "\Temp\CityName.txt"
        DownloadDVFile inurl, outpath
    '    pRS.MoveNext

End Sub
Private Sub DownloadDVFile(inurl As String, outpath As String)
   Set myie = CreateObject("InternetExplorer.Application")
 myie.Visible = True
 myie.Navigate URL:=inurl
 Dim WebString As String
 WebString = myie.Object.Document.DocumentElement.outerHTML
End Sub
Public Function GetZipFromCity(CityName As String, StateName As String) As String
    Dim ie     As Object
    Dim URL, ZipZ, SearchStr As String
    Dim WebString As String
    SearchStr = StateName & "-" & Replace(CityName, " ", "-")
    URL = "https://www.zip-codes.com/search.asp?srch-type=fuzzy&q=" & CityName & "%2C+" & StateName & ""
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = False
    ie.Navigate URL
    While ie.Busy
        DoEvents
    Wend
    WebString = ie.Document.DocumentElement.outerHTML
    ie.Quit
    Set ie = Nothing
    GetZipFromCity = SplitString(SplitString(SplitString(WebString, "onmouseover", (HowManyInString(WebString, SearchStr))), "/zip-code-", 0), "/zip-code/", 1)
    If GetZipFromCity Like "*Error:*" Then
        GetZipFromCity = "city not Found"
        MsgBox "There wasn't a match for the City: " & CityName & " in " & StateName & " State", vbCritical, Atlasversion
        Exit Function
    End If
    Dim intFileNum As Integer
    intFileNum = FreeFile
    ' change Output to Append if you want to add to an existing file
    ' rather than creating a new file each time
    Open "C:\Faas\" & FAUserName & "\Temp\WebText.txt" For Output As intFileNum
    Print #intFileNum, WebString
    Close intFileNum

End Function
Public Function LookUpByZipWeb(sZip As String, Optional ShowMessage As Boolean, _
                               Optional GetCity As Boolean, Optional GetState As Boolean, _
                               Optional GetCounty As Boolean, Optional GetFIPS As Boolean) As String
   LookUpByZipWeb = ""
   If Len(sZip) < 5 Then Exit Function
    Dim ie     As Object
    Dim URL    As String
    Dim WebString As String
    Call CloseOpenIEByURL("https://www.zip-codes.com/zip-code/")
    On Error GoTo FixMe
    If ShowMessage = True Then Call MessageUser("Open", "Getting Sales Tax Info", "Contacting Website...")
    Set ie = CreateObject("InternetExplorer.Application")
    URL = "https://www.zip-codes.com/zip-code/" & Left(sZip, 5) & "/zip-code-" & Left(sZip, 5) & ".asp"
    ie.Visible = False
    ie.Navigate URL
    While ie.Busy
        DoEvents
    Wend
    If ShowMessage = True Then Call MessageUser("Update", "Getting Sales Tax Info", "Reading data...")
    WebString = ie.Document.DocumentElement.outerHTML
    ie.Quit
    Set ie = Nothing
    Dim CityX, StateX, CountyX, AlaisX, FIPSx As String
    WebString = SplitString(SplitString(WebString, "ZIP Code " & Left(sZip, 5) & " Data</h2>", 1), "</table>", 0)
    CityX = SplitString(SplitString(SplitString(WebString, "City:", 1), "</A>", 0), ".asp" & Chr(34) & ">", 1)
    StateX = SplitString(SplitString(SplitString(WebString, "/state/", 1), " [", 0), ".asp" & Chr(34) & ">", 1)
    CountyX = ProperCase(SplitString(SplitString(SplitString(WebString, "/county/", 1), ",", 0), ".asp" & Chr(34) & ">", 1), 1)
    AlaisX = SplitString(SplitString(SplitString(WebString, "City Alias(es):", 1), "</TD></TR>", 0), "<TD class=info>", 1)
    FIPSx = SplitString(SplitString(WebString, "FIPS Code: ", 1), "" & Chr(34) & " href=", 0)
    If GetCity = True Then LookUpByZipWeb = LookUpByZipWeb & CityX & "|"
    If GetState = True Then LookUpByZipWeb = LookUpByZipWeb & StateX & "|"
    If GetCounty = True Then LookUpByZipWeb = LookUpByZipWeb & CountyX & "|"
    If GetFIPS = True Then LookUpByZipWeb = LookUpByZipWeb & FIPSx & "|"
    Exit Function
FixMe:
    On Error Resume Next
    Call CloseOpenIEByURL("https://www.zip-codes.com/zip-code/")
    ie.Quit
    Set ie = Nothing
    Call MessageUser("close")
    WebString = "<font size=3><b>Whoops, there was an error!</b></Font><br />Please try your request again..."
    Call Dialog.RichBox(WebString, vbCustom + vbQuestion, "Tax Lookup - Zip:" & sZip, , , 0, 1, 1, 1, "OK")
    On Error GoTo 0
End Function
Public Function LookUpTaxByZipWeb(sState, sZip As String, Optional sCity As String)
    Call CloseOpenIEByURL("http://www.salestaxstates.com")
    Dim ie     As Object
    Dim URL    As String
    Dim WebString As String
    Dim HeadString, TaxString As String
    Dim xStateCityZip As String
    Dim CountyX As String
        On Error GoTo FixMe
    Call MessageUser("Open", "Getting County Info", "Contacting Website...")
    CountyX = SplitString(LookUpByZipWeb(sZip, , , , True), "|", 0)
    Call MessageUser("Update", "Getting Sales Tax Info", "Contacting Website...")
    If Len(sState) <= 4 Then sState = DLookup("Statename", "TaxStates", "PostalAbbreviation='" & sState & "'")
    xStateCityZip = sState & "-" & sCity
    If sZip <> "" Then xStateCityZip = xStateCityZip & "-" & sZip
    Dim AliasX As String
    Dim ii     As Integer
    Set ie = CreateObject("InternetExplorer.Application")
    URL = "http://www.salestaxstates.com/sales-tax-calculator-" & xStateCityZip
    ie.Visible = False
    ie.Navigate URL
    While ie.Busy
        DoEvents
    Wend
    Call MessageUser("Update", "Getting Sales Tax Info", "Reading data...")
    WebString = ie.Document.DocumentElement.outerHTML
    ie.Quit
    Set ie = Nothing
    TaxString = SplitString(WebString, "onchange=" & Chr(34) & "javascript:calculSalesRates();" & Chr(34) & ">", 1)
    TaxString = SplitString(TaxString, "%", 0)
    TaxString = SplitString(TaxString, ">", 1) & "%"
    WebString = SplitString(WebString, "Sales taxes are calculated for zip code", 1)
    WebString = SplitString(WebString, "<h2>Last sales taxes rates update</h2>", 0)
    WebString = SplitString(WebString, "</h2>", 1)
    WebString = Replace(WebString, ",", "<br />", 1)
    WebString = Replace(WebString, "<br><br>", "<br />", 1)
    If WebString Like "*no sales tax for*" Then GoTo ShowMessage
    WebString = Replace(WebString, "The", "<font size=2>Zip:<b> ", 1, 1)
    WebString = Replace(WebString, "<br />", " </b>City:<b> ", 1, 1)
    WebString = Replace(WebString, "<br />", "<br /></b>State:<b> ", 1, 1)
    WebString = Replace(WebString, "general", "</b></font><br />The general ", 1, 1)
    WebString = Replace(WebString, "of the", "of:<br />", 1, 1)
    WebString = Replace(WebString, "<p class=" & Chr(34) & "text" & Chr(34) & ">", "<font size=1>", 1)
    WebString = Replace(WebString, "</p>", "</font>", 1)
    WebString = Replace(WebString, "rate is", "rate is<b>", 1)
    WebString = Replace(WebString, "%", "%</b>", 1)
    WebString = Replace(WebString, "rate (", "rate (<b>", 1)
    WebString = Replace(WebString, "%)", "%</b>)", 1)
    WebString = Replace(WebString, "<br /> the " & Left(sZip, 5) & "'s", "<br />" & CountyX, 1)
    WebString = SplitString(WebString, "Please refer", 0)
    WebString = WebString & "<br />To<b> Copy</b> or<b> Save</b> (as text), this info, use buttons below:<br />" & URL & ""

    HeadString = SplitString(WebString, "<br /> </b></font><br />", 0) & "</b>"
    If Len(CountyX) > 1 Then HeadString = HeadString & " County:<b> " & CountyX & "</b><br />"

    WebString = Replace(WebString, "the ", "", 1)
    WebString = HeadString & "Combined Sales Tax Rate:<b> " & TaxString & "</b></font><br />The " & SplitString(WebString, "%</b>.<br />", 1)
    WebString = Replace(WebString, "" & Left(sZip, 5) & "'s tax rate may change depending of type of purchase.", "<b>Note:</b> The county name listed above<i> may differ</i> from the county selected.", 1)
    WebString = Replace(WebString, " and in some case<br /> special rate", "and in some cases, a special rate of:", 1)
    If WebString Like "*Error*" Then
    WebString = "<font size=3><b>Error!</b></Font><br />There is no data matching your request<br />Requested: " & sState & " " & sZip
    End If

ShowMessage:
    Call MessageUser("close")
    Call Dialog.RichBox(WebString, vbCustom + vbInformation, "Tax Lookup - Zip:" & sZip, , , 0, 1, 1, 1, "OK")
Exit Function
FixMe:
On Error Resume Next
    Call CloseOpenIEByURL("http://www.salestaxstates.com")
    ie.Quit
    Set ie = Nothing
    Call MessageUser("close")
    WebString = "<font size=3><b>Whoops, there was an error!</b></Font><br />Please try your request again..."
    Call Dialog.RichBox(WebString, vbCustom + vbQuestion, "Tax Lookup - Zip:" & sZip, , , 0, 1, 1, 1, "OK")
On Error GoTo 0
    'http://www.salestaxstates.com/sales-tax-calculator-missouri-st_louis-63114
End Function
Function CloseOpenIEByURL(ByVal URL2Check As String) As SHDocVw.InternetExplorer
Dim objShellWindows As New SHDocVw.ShellWindows
'ignore errors when accessing the document property
On Error Resume Next
  'loop over all Shell-Windows
  For Each GetOpenIEByURL In objShellWindows
    'if the document is of type HTMLDocument, it is an IE window
    If TypeName(GetOpenIEByURL.Document) = "HTMLDocument" Then
      'check the URL
      If GetOpenIEByURL.Document.URL Like "*" & URL2Check & "*" Then
         GetOpenIEByURL.Quit
        'close, we found the right window
        Exit Function
      End If
    End If
  Next
End Function
Public Function LookUpZipWeb(Zipx As String, Optional Field As String) As String
    If Len(Zipx) < 4 Then Exit Function
    If Zipx = "city not Found" Then Exit Function
    Dim ie     As Object
    Dim URL    As String
    Dim WebString As String
    'Dim CityX, StateX, CountyX, FipsX As String
    Dim AliasX As String
    Dim ii     As Integer
    Dim TempDict As Object
    Set TempDict = New Collection
    Set ie = CreateObject("InternetExplorer.Application")
    URL = "https://www.zip-codes.com/zip-code/" & Left(Zipx, 5) & "/zip-code-" & Left(Zipx, 5) & ".asp"
    ie.Visible = False
    ie.Navigate URL
    While ie.Busy
        DoEvents
    Wend
    WebString = ie.Document.DocumentElement.outerHTML
    ie.Quit
    Set ie = Nothing
    Dim CityX, CountyX, StateX, FIPSx As String
    ' intFileNum = FreeFile
    ' change Output to Append if you want to add to an existing file
    ' rather than creating a new file each time

    '  Open "C:\Faas\" & FAUserName &"\Temp\WebText.txt" For Output As intFileNum
    '  Print #intFileNum, WebString
    ' Close intFileNum
    WebString = SplitString(SplitString(WebString, "ZIP Code " & Left(Zipx, 5) & " Data</h2>", 1), "</table>", 0)
    AliasX = SplitString(SplitString(WebString, "City Alias(es):", 1), "<tr><td class=", 0)
    AliasX = SplitString(SplitString(AliasX, "" & Chr(34) & "info" & Chr(34) & ">", 1), "</td></tr>", 0)
    CityX = SplitString(SplitString(SplitString(WebString, ">City:", 1), "</A>", 0), ".asp" & Chr(34) & ">", 1)
    StateX = SplitString(SplitString(SplitString(WebString, "/state/", 1), " [", 0), ".asp" & Chr(34) & ">", 1)
    CountyX = ProperCase(SplitString(SplitString(SplitString(WebString, "/county/", 1), ",", 0), ".asp" & Chr(34) & ">", 1), 1)
    FIPSx = SplitString(SplitString(WebString, "FIPS Code: ", 1), "" & Chr(34) & " href=", 0)
    'TempDict.Add Key:="AliasCities", Item:=HowManyInString(AliasX, "<br>")
    'If HowManyInString(AliasX, "<br>") > 0 Then
    '    For ii = 0 To HowManyInString(AliasX, "<br>")
    '        If SplitString(AliasX, "<br>", ii) <> TempDict("City") Then TempDict.Add Key:="AlaisCity" & (ii), Item:=SplitString(AliasX, "<br>", ii)
    '    Next ii
    'End If
    LookUpZipWeb = "|" & StateX & "|" & CountyX & "|" & CityX & "|" & FIPSx & "|"
    If Field = "State" Then LookUpZipWeb = StateX
    If Field = "County" Then LookUpZipWeb = CountyX
    If Field = "City" Then LookUpZipWeb = CityX
    If Field = "FIPS" Then LookUpZipWeb = FIPSx
End Function
Public Function HowManyInStr(ByRef LookIn As String, ByVal ToFind As String) As Integer
         Dim i As Long
         Dim xString As String
            For i = 1 To Len(LookIn)
                xString = Mid(LookIn, i, 1)
                If xString = ToFind Then HowManyInStr = HowManyInStr + 1
            Next i
End Function

Public Function HowManyInString(ByRef LookIn As String, ByVal ToFind As String) As Integer
         Dim xString As String
         Dim i As Integer
         Dim Plen As Integer
         Plen = Len(ToFind)
            For i = 1 To Len(LookIn)
                xString = Mid(LookIn, i, Plen)
                If xString = ToFind Then HowManyInString = HowManyInString + 1
            Next i
End Function
