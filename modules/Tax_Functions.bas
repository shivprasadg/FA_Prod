Attribute VB_Name = "Tax_Functions"
Option Compare Database
Option Explicit
Public msg1 As String

Public Function AskAboutTaxChange(TaxType As String, OldRate As Double, NewRate As Double) As String
Dim ResposeX As VbMsgBoxResultEx
msg1 = ""
msg1 = msg1 & "<b><Font Size=3>" & TaxType & " change...</font></b><br/>"
msg1 = msg1 & "From <b> " & Format(OldRate, "Percent") & "</b> to <b> " & Format(NewRate, "Percent") & "</b><br/><br/>"
msg1 = msg1 & "  <b>1:</b> This is a <b>Correction </b>, update this Tax Rate<br/>"
msg1 = msg1 & "  <b>2:</b> This is a <b>New Rate</b>, open the Add-A-Tax Wizard<br/>"
msg1 = msg1 & "  <b>3:</b> Neither, please <b>Undo</b> the change (default)<br/><br/> Please select..."
ResposeX = Dialog.RichBox(msg1, vbCustom + vbExclamation, AtlaasVersion, , , 0, 1, 0, 0, "Undo", "New Rate", "Correction")
If ResposeX = vbBt1 Then AskAboutTaxChange = "Undo"
If ResposeX = vbBt2 Then AskAboutTaxChange = "AddTax"
If ResposeX = vbBt3 Then AskAboutTaxChange = "Correction"
End Function
Public Function MakePercent(ValueX As String) As Double
If Nz(ValueX, "") = "" Then Exit Function
    If ValueX Like "*%*" Then ValueX = Replace(ValueX, "%", "")
    If ValueX > 0 And ValueX < 100 Then ValueX = CDbl(ValueX)
     If ValueX > 99 And ValueX < 1000 Then ValueX = CDbl(ValueX) / 10
      If ValueX >= 1000 Then ValueX = 0
    MakePercent = CDbl(ValueX)
End Function

Public Function isTaxActive(StartsOn As Date, EndsOn As Date) As Boolean
isTaxActive = False
If StartsOn > Date Then isTaxActive = False
If StartsOn <= Date And EndsOn <= Date Then isTaxActive = False
If StartsOn <= Date And EndsOn >= Date Then isTaxActive = True
End Function

Public Function ActiveTaxStatus(Taxlevelx As String, TaxTypeIDi As Integer, TaxLevelIDi As Integer) As String
    Dim SourceX As String
    Dim TaxFieldX As String
    Dim Namex  As String
    Dim LevelName As String
    Dim RST    As DAO.Recordset
    If Taxlevelx = "State" Then
        SourceX = "vw_Tax_StateBaseQ"
        TaxFieldX = "TaxStatesID"
        Namex = "StateName as NameX"
        LevelName = Nz(DLookup("State", "vw_CCSZ_Lookup", "StateID=" & TaxLevelIDi), "")

    End If
    msg1 = ""
    If Taxlevelx = "County" Then
        SourceX = "vw_Tax_CountyBaseQ"
        TaxFieldX = "TaxCountyID"
        Namex = "CountyName as NameX"
        LevelName = DLookup("County", "vw_CCSZ_Lookup", "CountyID=" & TaxLevelIDi)
    End If
    If Taxlevelx = "City" Then
        SourceX = "vw_Tax_CityBaseQ"
        TaxFieldX = "TaxCityID"
        Namex = "CityName as NameX"
        LevelName = DLookup("City", "vw_CCSZ_Lookup", "CityID=" & TaxLevelIDi)
    End If
    Set RST = CurrentDb.OpenRecordset( _
              "SELECT *," & Namex & " From " & SourceX & " WHERE TaxTypeID =" & TaxTypeIDi & " AND " & TaxFieldX & "=" & TaxLevelIDi & " AND IsActive = True" _
              , dbOpenSnapshot)
    If RST.RecordCount < 1 Then
        msg1 = msg1 & "<b><Font Size=3>No Active Tax!</font></b><br/>"
        msg1 = msg1 & "There isn't an Active<b> " & DLookup("TaxAbbreviation", "TaxType", "ID=" & TaxTypeIDi) & "</b> Tax<br/>"
        msg1 = msg1 & "for the " & Taxlevelx & " of:<b> " & LevelName & "</b><Br />"
        msg1 = msg1 & "Please Correct or Add in a Tax Rate with Proper Start and End Dates"
        GoTo FixMe
    End If

    If IsDate(RST!EndsOn) Then
        msg1 = msg1 & "<b><Font Size=3>Active Tax Set to Expire!</font></b><br/>"
        msg1 = msg1 & RST!Namex & "'s Active " & RST!TaxType & " Tax Rate of " & Format(RST!BaseTaxRate, "##0.###") & "% has a start date of " & RST!StartsOn & "<br />"
        msg1 = msg1 & "with an <b>expiration date of " & RST!EndsOn & "</b><br />"
        msg1 = msg1 & "Please<b> validate </b>the expiration date."
    End If

    If Not IsDate(RST!EndsOn) Then
        msg1 = msg1 & "<b><Font Size=3>" & RST!TaxType & " Tax is Valid</font></b><br/>"
        msg1 = msg1 & RST!Namex & "'s Active<b> " & RST!TaxType & "</b> Tax Rate of<b> " & Format(RST!BaseTaxRate, "##0.###") & "%</b><br />"
        msg1 = msg1 & "has a start date of<b> " & RST!StartsOn & "</b><br /> and will continue in<b> perpetuity</b>"
    End If
FixMe:
ActiveTaxStatus = msg1
    RST.close
    Set RST = Nothing
End Function
Public Function TaxStatus(Taxlevelx As String, TaxTypeIDi As Integer, TaxLevelIDi As Integer) As String
    Dim SourceX As String
    Dim TaxFieldX As String
    Dim Namex  As String
    Dim LevelName As String
    Dim SQLi   As String
    Dim ii     As Integer
    Dim ThereIsATax As Boolean
    Dim IsFutureTax As Boolean
    Dim RST    As DAO.Recordset
    If Taxlevelx = "State" Then
        SourceX = "vw_Tax_StateBaseQ"
        TaxFieldX = "TaxStatesID"
        Namex = "StateName as NameX"
        LevelName = Nz(DLookup("State", "vw_CCSZ_Lookup", "StateID=" & TaxLevelIDi), "")

    End If
    msg1 = ""
    If Taxlevelx = "County" Then
        SourceX = "vw_Tax_CountyBaseQ"
        TaxFieldX = "TaxCountyID"
        Namex = "CountyName as NameX"
        LevelName = DLookup("County", "vw_CCSZ_Lookup", "CountyID=" & TaxLevelIDi)
    End If
    If Taxlevelx = "City" Then
        SourceX = "vw_Tax_CityBaseQ"
        TaxFieldX = "TaxCityID"
        Namex = "CityName as NameX"
        LevelName = DLookup("City", "vw_CCSZ_Lookup", "CityID=" & TaxLevelIDi)
    End If
    SQLi = SQLi & "SELECT *," & Namex & " From " & SourceX & " "
    SQLi = SQLi & "WHERE TaxTypeID =" & TaxTypeIDi & " "
    SQLi = SQLi & "AND " & TaxFieldX & "=" & TaxLevelIDi & " "
    SQLi = SQLi & "Order By Nz(EndsOn, '12/31/2050') DESC;"
    Set RST = CurrentDb.OpenRecordset(SQLi, dbOpenSnapshot)
    If RST.RecordCount > 0 Then
        ThereIsATax = True
        RST.MoveLast
        RST.MoveFirst
        Dim StartsOnx As Date
        Dim EndsOnx As Date
        For ii = RST.RecordCount To 1 Step -1
            EndsOnx = Nz(RST!EndsOn, "12/31/2050")
            StartsOnx = RST!StartsOn
            If RST!StartsOn > Date And Nz(RST!EndsOn, "12/31/2050") > Date Then IsFutureTax = True
            If ii < RST.RecordCount Then
                RST.MovePrevious
                If StartsOnx > RST!StartsOn And EndsOnx < Nz(RST!EndsOn, "12/31/2050") Then TaxStatus = "OverLap"
                RST.MoveNext
            End If
            RST.MoveNext
        Next ii
    Else
        ThereIsATax = False
    End If
    SQLi = ""
    SQLi = SQLi & "SELECT *," & Namex & " From " & SourceX & " "
    SQLi = SQLi & "WHERE TaxTypeID =" & TaxTypeIDi & " "
    SQLi = SQLi & "AND " & TaxFieldX & "=" & TaxLevelIDi & " "
    SQLi = SQLi & "AND IsActive = True"

    Set RST = CurrentDb.OpenRecordset(SQLi, dbOpenSnapshot)
    If RST.RecordCount < 1 Then
        If ThereIsATax = True And TaxStatus <> "Overlap" Then TaxStatus = "No Active"
        If ThereIsATax = False And TaxStatus <> "Overlap" Then TaxStatus = "No Tax"
        GoTo FixMe
    End If
    If IsDate(RST!EndsOn) And TaxStatus <> "Overlap" Then TaxStatus = "Will Expire"
    If Not IsDate(RST!EndsOn) And TaxStatus <> "Overlap" Then TaxStatus = "Valid"
    If TaxStatus = "Will Expire" And IsFutureTax = True Then TaxStatus = "Valid w/Future"
FixMe:
   ' TaxStatus = TaxStatus
    RST.close
    Set RST = Nothing
End Function
Public Function RefreshTaxForms()
    If IsOpen("ClientsView") Then
        Dim TempLocID As Long
        TempLocID = Forms!ClientsView!Locations_sub!LocationList
        Forms!ClientsView!Locations_sub!LocationTax.Requery
        Forms!ClientsView!Locations_sub.Requery
        Forms!ClientsView!Locations_sub!LocationList = TempLocID
        Form_Locations_sub.LocationList_Click
    End If
    If IsOpen("Tax_State") Then Forms!Tax_State!Tax_StateBaseSub.Requery
    If IsOpen("Tax_County") Then Forms!Tax_County!Tax_CountyBaseSub.Requery
    If IsOpen("Tax_City") Then Forms!Tax_City!Tax_CityBaseSub.Requery
    If IsOpen("LocationAddressesEdit") Then Forms!LocationAddressesEdit!LocationAddressList.Requery
End Function
Public Function GetNADA(Optional VinX As String) As String
    Dim eItem  As IHTMLElement
    Dim oBrowser As InternetExplorer
    Dim ie     As Variant
    Dim URL    As String
    Dim ii     As Integer
    Dim WebString As String
    Dim SQLz   As String
    Dim Tempx  As String
    Dim Wholesale, Loan, Retail, MSRP As Currency
    Dim RST    As DAO.Recordset
    Set RST = CurrentDb.OpenRecordset("Select * From NADAResults WHERE nz(ReturnedResults,0)=0", dbOpenSnapshot)
    RST.MoveLast
    RST.MoveFirst
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Silent = True
    ie.Visible = True    ' set to false to hide true to show
    ' Call MessageUser("Open", "Contacting Zip2Tax.com", "Opening web connection...")
    ' Your webpage goes here
    URL = "https://extapps.nada.com/NADAOnline/ct/"
    ie.Navigate URL
    While ie.ReadyState <> 4: DoEvents: Wend

    For ii = 1 To RST.RecordCount
        VinX = RST!PilotVIN
        If ii = 1 Then
            For Each eItem In ie.Document.getElementsByTagName("input")
                If eItem.ID = "usernameInput" Then eItem.Value = "fleetadvantage"
                If eItem.ID = "passwordInput" Then eItem.Value = "fleetadv"
            Next
            Do While ie.Busy Or ie.ReadyState <> 4
                DoEvents
            Loop

            'Call MessageUser("Update", "Submitting Request", "Sending data...")
            For Each eItem In ie.Document.getElementsByTagName("button")
                If eItem.ID = "loginButton" Then
                    eItem.Click
                    Exit For
                End If
            Next
        End If
        Do While ie.Busy Or ie.ReadyState <> 4
            DoEvents
        Loop
        'Do Until ie.ReadyState = 4: DoEvents: Loop
        Pause (3)
        For Each eItem In ie.Document.getElementsByTagName("input")
            If eItem.ID = "vinDecodeTextBox" Then
                eItem.Focus
                eItem.Value = VinX
                Exit For
            End If
        Next
        Do While ie.Busy Or ie.ReadyState <> 4
            DoEvents
        Loop
        ' While ie.ReadyState <> READYSTATE_COMPLETE And ie.ReadyState <> READYSTATE_LOADED
        '    DoEvents
        ' Wend
        Pause (3)
        For Each eItem In ie.Document.getElementsByTagName("button")
            If eItem.ID = "vinDecodeButton" Then
                eItem.removeAttribute ("disabled")
                eItem.Click
                Exit For
            End If
        Next
        Do While ie.Busy Or ie.ReadyState <> 4
            DoEvents
        Loop
        Pause (2)
        Wholesale = 0
        Loan = 0
        Retail = 0
        MSRP = 0
        WebString = ie.Document.DocumentElement.outerHTML

        If WebString Like "*<h1>Multiple VIN Results:</h1>*" Then
            ie.Document.parentWindow.execScript "multipleVinResultSelected();", "javascript"
            Do While ie.Busy Or ie.ReadyState <> 4
                DoEvents
            Loop
            Pause (2)
             WebString = ie.Document.DocumentElement.outerHTML
        End If

        If WebString Like "*Invalid VIN. Please re-enter.*" Then
            For Each eItem In ie.Document.getElementsByTagName("button")
                If eItem.outerText = "OK" Then
                    WebString = ""
                    eItem.Click
                    ie.Navigate URL
                    Do While ie.Busy Or ie.ReadyState <> 4
                        DoEvents
                    Loop
                    WebString = ie.Document.DocumentElement.outerHTML
                    Exit For
                End If
            Next
        End If
         Pause (3)
        Tempx = SplitString(SplitString(WebString, "baseMSRP", 1), "</td>", 0)
        Tempx = Replace(Tempx, Chr(34) & ">$", "")
        Tempx = Replace(Tempx, Chr(34) & ",", "")
        If Not IsNumeric(Tempx) Then MSRP = 0 Else MSRP = Tempx
        '<button class="pull-right" onclick="return multipleVinResultSelected();">Select</button>
        Tempx = SplitString(SplitString(WebString, "tblVehicleValues3", 1), "</tbody>", 0)
        Tempx = SplitString(Tempx, "Base</label></th>", 1)
        Tempx = SplitString(Tempx, "</tr>", 0)
        Tempx = Trim(Tempx)
        Tempx = Replace(Tempx, Chr(34) & ">$", "|")
        Tempx = Replace(Tempx, "<td class=" & Chr(34), "")
        Tempx = Replace(Tempx, "</td>", ":")
        Tempx = Replace(Tempx, ",", "")
        Tempx = Replace(Tempx, "|", ",")
        If IsNumeric(SplitString(SplitString(Tempx, "Wholesale,", "1"), ":", 0)) Then Wholesale = SplitString(SplitString(Tempx, "Wholesale,", "1"), ":", 0)
        If IsNumeric(SplitString(SplitString(Tempx, "Retail,", "1"), ":", 0)) Then Retail = SplitString(SplitString(Tempx, "Retail,", "1"), ":", 0)
        If IsNumeric(SplitString(SplitString(Tempx, "Loan,", "1"), ":", 0)) Then Loan = SplitString(SplitString(Tempx, "Loan,", "1"), ":", 0)
        SQLz = "Update NADAResults Set "
        SQLz = SQLz & "WholeSaleBase = " & Wholesale & ""
        SQLz = SQLz & ",RetailBase =" & Retail & ""
        SQLz = SQLz & ",LoanBase =" & Loan & ""
        SQLz = SQLz & ",MSRP =" & MSRP & ""
        If MSRP > 0 Then SQLz = SQLz & ",ReturnedResults =-1" Else SQLz = SQLz & ",ReturnedResults =0"
        SQLz = SQLz & ",DateRequested='" & Date & "' "
        SQLz = SQLz & " WHERE PilotVIN='" & VinX & "'"

        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
        RST.MoveNext
        WebString = ""
GetNew:
        For Each eItem In ie.Document.getElementsByTagName("button")
            If eItem.ID = "btnResetVehicle" Then
                'eItem.removeAttribute ("disabled")
                eItem.Click
                Exit For
            End If
        Next
        Do While ie.Busy Or ie.ReadyState <> 4
            DoEvents
        Loop
        'End If
    Next ii
    '   Wholesale = SplitString(TempX, "Wholesale,", 1)
    ie.Quit
    Set ie = Nothing
    RST.close
    Set RST = Nothing
End Function
OpenEdge:
Public Function LookupSaleTaxStates(Zipx As String, Optional CurrentTax As Double, Optional tStateID As Long, Optional tCountyID As Long, _
    Optional tCityID As Long, Optional tLocationID As Long, Optional CountyName As String) As String
    Dim wb As WebDriver
    Dim URL, Tempx, Place As String
    Dim Headerx As String
    Dim nCounty, nCity, nState As String
    Dim SQLt, SQLz As String
    Dim FullTax As Double
    Dim CountyTax As Double
    Dim StateTax As Double
    Dim CityTax As Double
    Dim SpecialTax As Double
    Dim ResponseX As Integer
    Call MessageUser("Open", "Contacting SalesTaxStates.com", "Opening web connection...")
    Set wb = New Selenium.EdgeDriver     ' Show Browser
    'Set WB = New Selenium.PhantomJSDriver ' Hidden
    URL = "https://www.salestaxstates.com"
    wb.start
    wb.Get URL
    wb.FindElementById("autocomplete").SendKeys Zipx
    Call MessageUser("Update", "Posting Request to SalesTaxStates", "Sending ZipCode: ( " & Zipx & " )")
    wb.Wait (1000)
    wb.FindElementById("goSubmit").Click
    Call MessageUser("Update", "Waiting for Response", "Getting Taxes for: ( " & Zipx & " )")
    wb.Wait (1000)
    Place = wb.FindElementByCss("title").Attribute("outerHTML")
    nCity = SplitString(SplitString(Place, Zipx, 1), ",", 0)
    nState = SplitString(SplitString(Place, nCity & ",", 1), ",", 0)
    nCounty = CountyName
    Tempx = wb.FindElementByXPath("/html/body/section/section[1]/section[3]/section/form/p[1]/span[2]").Attribute("outerHTML")
    FullTax = CDbl(wb.FindElementByXPath("/html/body/section/section[1]/section[3]/section/form/p[1]/span[2]/select/option[1]").Value)
    If Tempx Like "*(State)*" Then StateTax = CDbl(wb.FindElementByXPath("/html/body/section/section[1]/section[3]/section/form/p[1]/span[2]/select/option[2]").Value)
    If Tempx Like "*(County)*" Then CountyTax = CDbl(wb.FindElementByXPath("/html/body/section/section[1]/section[3]/section/form/p[1]/span[2]/select/option[3]").Value)
    If Tempx Like "*(City)*" And CountyTax = 0 Then CityTax = CDbl(wb.FindElementByXPath("/html/body/section/section[1]/section[3]/section/form/p[1]/span[2]/select/option[3]").Value)
    If Tempx Like "*(City)*" And CountyTax > 0 Then CityTax = CDbl(wb.FindElementByXPath("/html/body/section/section[1]/section[3]/section/form/p[1]/span[2]/select/option[4]").Value)
    If Tempx Like "*(Special)*" And CityTax = 0 Then SpecialTax = CDbl(wb.FindElementByXPath("/html/body/section/section[1]/section[3]/section/form/p[1]/span[2]/select/option[4]").Value)
    If Tempx Like "*(Special)*" And CountyTax = 0 Then SpecialTax = CDbl(wb.FindElementByXPath("/html/body/section/section[1]/section[3]/section/form/p[1]/span[2]/select/option[4]").Value)
    If Tempx Like "*(Special)*" And CountyTax > 0 And CityTax > 0 Then SpecialTax = CDbl(wb.FindElementByXPath("/html/body/section/section[1]/section[3]/section/form/p[1]/span[2]/select/option[5]").Value)


    If FullTax > 0 And StateTax + CountyTax + CityTax = 0 Then StateTax = FullTax
    Call MessageUser("Close")
    Headerx = "<font size=3>Sales Tax<b> " & Format(FullTax, "#.###%") & "</b></font><br />( for zip code: " & Zipx & " )<br /><br />"
    LookupSaleTaxStates = ""
    LookupSaleTaxStates = LookupSaleTaxStates & "<b>Details:</b><br />"
    LookupSaleTaxStates = LookupSaleTaxStates & "State of<b> " & nState & "</b> : " & Format(StateTax, "#.###%") & "<br />"
    LookupSaleTaxStates = LookupSaleTaxStates & "County of<b> " & nCounty & "</b> : " & Format(CountyTax, "#.###%") & "<br />"
    LookupSaleTaxStates = LookupSaleTaxStates & "City of<b> " & nCity & "</b> : " & Format(CityTax, "#.###%") & "<br />"
    If SpecialTax > 0 Then LookupSaleTaxStates = LookupSaleTaxStates & "Plus a<b> Special</b> Tax : " & Format(SpecialTax, "#.###%") & "<br />"
    LookupSaleTaxStates = LookupSaleTaxStates & "<br />Tax Rate Information retrived from<br />SaleTaxStates.com<br />"
    wb.close
    If FullTax <> CurrentTax Then
        Headerx = "<font size=3>Taxes Don't Match!<br />"
        Headerx = Headerx & "The Rate Is:<b> " & Format(FullTax, "Percent") & "</b></font><br />for zip code: " & Zipx & "<br /><br />"
        Headerx = Headerx & "<font size=2>Atlaas Core has:<b> " & Format(CurrentTax, "Percent") & "</b> which may be wrong.<br /></font><br />"
        LookupSaleTaxStates = Headerx & LookupSaleTaxStates & "<br /><i>To<b> Copy</b> or<b> Save</b> (as text), use 'copy/disk' buttons below</i><br /><br />"
        LookupSaleTaxStates = LookupSaleTaxStates & "<font size=2>To<b> Correct</b> the Taxes, press <b><i>Correct It</i></b></font>"
        ResponseX = Dialog.RichBox(LookupSaleTaxStates, vbCustom + vbCritical, "The Stored Tax Rate Dosen't Match!", , , 0, 1, 1, 1, "Leave it Alone", "Correct It")
        '101=Leave it Alone
        '102=Correct Taxes
        If ResponseX = 102 And tLocationID > 0 Then  'Correct the Taxes
            Dim RST    As DAO.Recordset
            SQLt = "Update Locations Set LTaxRate = " & FullTax & " WHERE LocationID=" & tLocationID & ";"
            Call MessageUser("Open", "Updating Location Tax to: " & FullTax & "", "Running SQL Command")
            DoCmd.SetWarnings False
            DoCmd.RunSQL (SQLt)
            DoCmd.SetWarnings True
        End If
    Else
        LookupSaleTaxStates = Headerx & LookupSaleTaxStates & "<br /><i>To<b> Copy</b> or<b> Save</b> (as text), use 'copy/disk' buttons below</i><br /><br />"
        Call Dialog.RichBox(LookupSaleTaxStates, vbCustom + vbInformation, "Tax Lookup - Zip: " & Zipx, , , 0, 1, 1, 1, "OK")
    End If
End Function

Public Function CleanJSON(s As String) As String
    s = Replace(s, """", "")
    s = Replace(s, "\n", vbCrLf)
    CleanJSON = s
End Function
'Public Sub Pause(ms As Long)
'    Dim t As Single
'    t = Timer
'    Do While Timer < t + (ms / 1000)
'        DoEvents
'    Loop
'End Sub
Public Function LookUpTaxByZip2Tax_zXX1(Zipx As String, Optional CurrentTax As Double, Optional tStateID As Long, Optional tCountyID As Long, _
                                                   Optional tCityID As Long, Optional tLocationID As Long) As String
    Dim wb As WebDriver
    Dim URL, Tempx, Place As String
    Dim FullTax As String
    Dim nState, StateTax As String
    Dim nCounty, CountyTax As String
    Dim nCity, CityTax As String
    Dim SQLt, SQLz As String

    Call MessageUser("Open", "Contacting Zip2Tax.com", "Opening web connection...")
    Set wb = New Selenium.EdgeDriver
    'WB.SetCapability "ms:edgeOptions", "{""args"":[""--headless""]}"
    'Set WB = New Selenium.PhantomJSDriver
    URL = "https://myaccount.zip2tax.com/user/login"
    wb.start
    wb.Get URL
    wb.FindElementById("Email").SendKeys "fleeta"
    wb.FindElementById("Password").SendKeys "af~52fn"
    wb.FindElementByXPath("/html/body/div[1]/div[1]/div[2]/div/div/div/div/form/div[5]/div/button").Click
    wb.Wait (1000)
    wb.FindElementById("txtZip").SendKeys Zipx
    wb.FindElementByXPath("/html/body/div[1]/div[1]/div[2]/div/form/div/div[2]/div[1]/button").Click
    wb.Wait (1000)
    wb.FindElementByXPath("/html/body/div[1]/div[1]/div[2]/div/div[2]/div[3]/div/h2/button").Click
    wb.Wait (1000)
    FullTax = wb.FindElementByXPath("/html/body/div[1]/div[1]/div[2]/div/div[2]/div[2]/div/div[1]/div/div[1]/div[2]").Text
    Place = wb.FindElementByXPath("/html/body/div[1]/div[1]/div[2]/div/div[2]/div[2]/div/div[1]/div/div[2]/div[5]").Text
    nState = wb.FindElementByXPath("/html/body/div[1]/div[1]/div[2]/div/div[2]/div[3]/div/div/div/div/div[2]/div[1]").Text
    StateTax = wb.FindElementByXPath("/html/body/div[1]/div[1]/div[2]/div/div[2]/div[3]/div/div/div/div/div[2]/div[2]").Text
    nCounty = wb.FindElementByXPath("/html/body/div[1]/div[1]/div[2]/div/div[2]/div[3]/div/div/div/div/div[3]/div[1]").Text
    CountyTax = wb.FindElementByXPath("/html/body/div[1]/div[1]/div[2]/div/div[2]/div[3]/div/div/div/div/div[3]/div[2]").Text
    On Error Resume Next
    nCity = wb.FindElementByXPath("/html/body/div[1]/div[1]/div[2]/div/div[2]/div[3]/div/div/div/div/div[4]/div[1]").Text
    CityTax = wb.FindElementByXPath("/html/body/div[1]/div[1]/div[2]/div/div[2]/div[3]/div/div/div/div/div[4]/div[2]").Text
    On Error GoTo 0
    If nCity Like "*Total*" Then
        nCity = "City of " & Place
        CityTax = "0%"
    End If
    If nCity = "" And Place <> "" Then nCity = "City of " & Place
    If CityTax = "" Then CityTax = "0%"
    If nCity Like "*Special*" Then nCity = "City of " & Place & " (<i>special</i>)"
    LookUpTaxByZip2Tax = "<font size=3>Sales Tax<b> " & FullTax & "</b></font><br />( for zip code " & Zipx & " )<br /><br />Details:<br />" & nState & " = " & StateTax & "<br />" & nCounty & " = " & CountyTax & "<br />" & nCity & " = " & CityTax & ""
    LookUpTaxByZip2Tax = LookUpTaxByZip2Tax & "<br /><br />Tax Rate Information retrived from Zip2Tax.com<br />"
    wb.close
    Dim iFullTax As Double
    Dim ResponseX As Integer

    iFullTax = CDbl(SplitString(FullTax, "%", 0)) / 100
    If iFullTax <> CurrentTax Then
        LookUpTaxByZip2Tax = LookUpTaxByZip2Tax & "<br /><B>The Taxes Don't Match!</B>"
        LookUpTaxByZip2Tax = LookUpTaxByZip2Tax & "<br />To<b> Copy</b> or<b> Save</b> (as text), use buttons below"
        ResponseX = Dialog.RichBox(LookUpTaxByZip2Tax, vbCustom + vbCritical, "The Stored Tax Rate Dosen't Match!", , , 0, 1, 1, 1, "Leave It", "Correct It")
        '101=Leave it Alone
        '102=Correct Taxes
        If ResponseX = 102 And tLocationID > 0 Then  'Correct the Taxes
           ' Dim RST    As DAO.Recordset
            SQLt = "Update Locations Set LTaxRate = " & iFullTax & " WHERE LocationID=" & tLocationID & ";"
            Call MessageUser("Open", "Updating Location Tax to: " & FullTax & "", "Running SQL Command")
            DoCmd.SetWarnings False
            DoCmd.RunSQL (SQLt)
            DoCmd.SetWarnings True
        End If
    Else
        LookUpTaxByZip2Tax = LookUpTaxByZip2Tax & "To<b> Copy</b> or<b> Save</b> (as text), use buttons below"
        Call Dialog.RichBox(LookUpTaxByZip2Tax, vbCustom + vbInformation, "Tax Lookup - Zip:" & Zipx, , , 0, 1, 1, 1, "OK")
    End If
    Call MessageUser("Close")
End Function


Public Function LookUpTaxByZip2Tax(Zipx As String, Optional CurrentTax As Double, Optional tStateID As Long, Optional tCountyID As Long, _
                                                   Optional tCityID As Long, Optional tLocationID As Long) As String
    Dim wb As WebDriver
    Dim URL, Tempx, Place As String
    Dim FullTax As String
    Dim nState, StateTax As String
    Dim nCounty, CountyTax As String
    Dim nCity, CityTax As String
    Dim nSpecial, SpecialTax
    Dim SQLt, SQLz As String
    Dim msg As String

    Call MessageUser("Open", "Contacting Zip2Tax.com", "Opening web connection...")
    Set wb = New Selenium.EdgeDriver

    URL = "https://myaccount.zip2tax.com/user/login"
    wb.start
    wb.Get URL
    wb.FindElementById("Email").SendKeys "fleeta"
    wb.FindElementById("Password").SendKeys "af~52fn"
    wb.FindElementByXPath("/html/body/div[1]/div[1]/div[2]/div/div/div/div/form/div[5]/div/button").Click
    wb.Wait (1000)
    wb.FindElementById("txtZip").SendKeys Zipx
    wb.FindElementByXPath("/html/body/div[1]/div[1]/div[2]/div/form/div/div[2]/div[1]/button").Click
    wb.Wait (1000)
   wb.FindElementByXPath("//*[@id=""results""]/div[3]/div/h2/button").Click
   ' wb.FindElementByXPath("/html/body/div[1]/div[1]/div[2]/div/div[2]/div[3]/div/h2/button").Click
    wb.Wait (1000)
    FullTax = wb.FindElementByXPath("/html/body/div[1]/div[1]/div[2]/div/div[3]/div[3]/div/div/div/div/div[5]/div[2]").Text
    Place = wb.FindElementByXPath("/html/body/div[1]/div[1]/div[2]/div/div[3]/div[2]/div/div[1]/div/div[2]/div[3]").Text
    nState = wb.FindElementByXPath("/html/body/div[1]/div[1]/div[2]/div/div[3]/div[2]/div/div[1]/div/div[2]/div[5]").Text
    StateTax = wb.FindElementByXPath("/html/body/div[1]/div[1]/div[2]/div/div[3]/div[3]/div/div/div/div/div[2]/div[2]").Text
    nCounty = wb.FindElementByXPath("/html/body/div[1]/div[1]/div[2]/div/div[3]/div[2]/div/div[1]/div/div[2]/div[9]").Text
    CountyTax = wb.FindElementByXPath("/html/body/div[1]/div[1]/div[2]/div/div[3]/div[3]/div/div/div/div/div[4]/div[2]").Text
    On Error Resume Next
    nSpecial = wb.FindElementByXPath("/html/body/div[1]/div[1]/div[2]/div/div[3]/div[3]/div/div/div/div/div[4]/div[1]").Text
    If nSpecial Like "*special*" Then SpecialTax = wb.FindElementByXPath("/html/body/div[1]/div[1]/div[2]/div/div[3]/div[3]/div/div/div/div/div[4]/div[2]").Text
    nCity = wb.FindElementByXPath("/html/body/div[1]/div[1]/div[2]/div/div[3]/div[3]/div/div/div/div/div[3]/div[1]").Text
    CityTax = wb.FindElementByXPath("/html/body/div[1]/div[1]/div[2]/div/div[3]/div[3]/div/div/div/div/div[3]/div[2]").Text
    On Error GoTo 0
    If nCity Like "*Total*" Then
        nCity = "City of " & Place
        CityTax = "0%"
    End If
    If nCity = "" And Place <> "" Then nCity = "City of " & Place
    If CityTax = "" Then CityTax = "0%"
    If nCity Like "*Special*" Then nCity = "City of " & Place & " (<i>special</i>)"
    LookUpTaxByZip2Tax = "<font size=3>Sales Tax<b> " & FullTax & "</b></font><br />( for zip code " & Zipx & " )<br /><br />Details:<br />" & nState & " = " & StateTax & "<br />" & nCounty & " = " & CountyTax & "<br />" & nCity & " = " & CityTax & ""
    LookUpTaxByZip2Tax = LookUpTaxByZip2Tax & "<br /><br />Tax Rate Information retrived from Zip2Tax.com<br />"
    wb.close
    Dim iFullTax As Double
    Dim ResponseX As Integer

    iFullTax = CDbl(SplitString(FullTax, "%", 0)) / 100
    If iFullTax <> CurrentTax Then
        LookUpTaxByZip2Tax = LookUpTaxByZip2Tax & "<br /><B>The Taxes Don't Match!</B>"
        LookUpTaxByZip2Tax = LookUpTaxByZip2Tax & "<br />To<b> Copy</b> or<b> Save</b> (as text), use buttons below"
        ResponseX = Dialog.RichBox(LookUpTaxByZip2Tax, vbCustom + vbCritical, "The Stored Tax Rate Dosen't Match!", , , 0, 1, 1, 1, "Leave It", "Correct It")
        '101=Leave it Alone
        '102=Correct Taxes
        If ResponseX = 102 And tLocationID > 0 Then  'Correct the Taxes
           ' Dim RST    As DAO.Recordset
            SQLt = "Update Locations Set LTaxRate = " & iFullTax & " WHERE LocationID=" & tLocationID & ";"
            Call MessageUser("Open", "Updating Location Tax to: " & FullTax & "", "Running SQL Command")
            DoCmd.SetWarnings False
            DoCmd.RunSQL (SQLt)
            DoCmd.SetWarnings True

        End If
    Else
        LookUpTaxByZip2Tax = LookUpTaxByZip2Tax & "To<b> Copy</b> or<b> Save</b> (as text), use buttons below"
        Call Dialog.RichBox(LookUpTaxByZip2Tax, vbCustom + vbInformation, "Tax Lookup - Zip:" & Zipx, , , 0, 1, 1, 1, "OK")
    End If
    Call MessageUser("Close")
End Function

Public Function OLDLookUpTaxByZip2Tax(Zipx As String) As String
    Dim oHTML_Element As IHTMLElement
    Dim oBrowser As InternetExplorer
    Dim ie     As Variant
    Dim URL As String
    Dim WebString As String
    Dim StateX As String
    Dim CityX  As String
    Dim CountyX As String
    Dim Tempx  As String
    Dim TempRate As Double
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Silent = True
    ie.Visible = True ' set to false to hide true to show
    Call MessageUser("Open", "Contacting Zip2Tax.com", "Opening web connection...")
    ' ie.Navigate "https://www.zip2tax.com/"    ' Your webpage goes here
    URL = "https://www.zip2tax.com/Website/pagesTaxRates/z2t_lookup.asp?inputZip=" & Zipx & ""
    ie.Navigate URL
    While ie.ReadyState <> READYSTATE_COMPLETE And ie.ReadyState <> READYSTATE_LOADED
        DoEvents
    Wend
    Call MessageUser("Update", "Logging in", "Submitting login info...")
    For Each oHTML_Element In ie.Document.getElementsByTagName("a")
        If oHTML_Element.ClassName = "dropdown-toggle login_menu" Then
            oHTML_Element.Click
        End If
    Next
    For Each oHTML_Element In ie.Document.getElementsByTagName("input")
        If oHTML_Element.ID = "ftUsername" Then oHTML_Element.Value = "fleeta"
        If oHTML_Element.ID = "ftPassword" Then oHTML_Element.Value = "af~52fn"
    Next
    Call MessageUser("Update", "Submitting Request", "Sending data...")
    For Each oHTML_Element In ie.Document.getElementsByTagName("button")
        If oHTML_Element.OnClick Like "*ftLoginNew*" Then oHTML_Element.Click
    Next
    While ie.ReadyState <> READYSTATE_COMPLETE And ie.ReadyState <> READYSTATE_LOADED
        DoEvents
    Wend
    '  For Each oHTML_Element In ie.Document.getElementsByTagName("input")
    '      If oHTML_Element.ID = "menuInputZip" Then
    '      oHTML_Element.Value = Zipx
    '      oHTML_Element.Enter
    '      End If
    '  Next
    Call MessageUser("Update", "Getting Sales Tax Info", "Reading data...")
    ie.Navigate URL   'ZipCode's tax
    While ie.ReadyState <> READYSTATE_COMPLETE And ie.ReadyState <> READYSTATE_LOADED
        DoEvents
    Wend
    WebString = ie.Document.DocumentElement.outerHTML
    WebString = SplitString(WebString, "Sales Tax Breakout For ", 1)
    ie.Quit
    Set ie = Nothing
    CityX = SplitString(WebString, "City:", 1)
    CityX = SplitString(CityX, "&nbsp;", 0)
    CityX = Trim(SplitString(CityX, ">", 2))
    CityX = "City of " & CleanTrim(CityX)

    StateX = SplitString(WebString, "result_state", 1)
    StateX = SplitString(StateX, "</span>", 0)
    StateX = Trim(SplitString(StateX, ">", 1))
    StateX = "State of " & CleanTrim(StateX)

    CountyX = SplitString(WebString, "County:", 1)
    CountyX = SplitString(CountyX, "&nbsp;", 0)
    CountyX = Trim(SplitString(CountyX, ">", 2))
    CountyX = "County of " & CleanTrim(CountyX)
    WebString = SplitString(WebString, "</TD></TR></TABLE>", 0) & "</TD></TR></TABLE>"

    If InStr(WebString, "City of") > 0 Then
        Tempx = SplitString(WebString, "City of", 1)
        Tempx = SplitString(Tempx, ">", 2)
        Tempx = SplitString(Tempx, "<", 0)
        CityX = CityX & "<b> " & Tempx & "</b>"
        TempRate = CDbl(SplitString(Tempx, "%", 0))
    Else
        CityX = CityX & " 0%</b>"
        TempRate = 0
    End If
    If InStr(WebString, "State of") > 0 Then
        Tempx = SplitString(WebString, "State of", 1)
        Tempx = SplitString(Tempx, ">", 2)
        Tempx = SplitString(Tempx, "<", 0)
        StateX = StateX & "<b> " & Tempx & "</b>"
        TempRate = TempRate + CDbl(SplitString(Tempx, "%", 0))
    Else
        StateX = StateX & " 0%</b>"
        TempRate = TempRate + 0
    End If
    If InStr(WebString, "County of") > 0 Then
        Tempx = SplitString(WebString, "County of", 1)
        Tempx = SplitString(Tempx, ">", 2)
        Tempx = SplitString(Tempx, "<", 0)
        CountyX = CountyX & "<b> " & Tempx & "</b>"
        TempRate = TempRate + CDbl(SplitString(Tempx, "%", 0))
    Else
        CountyX = CountyX & " 0%</b>"
        TempRate = TempRate + 0
    End If
    LookUpTaxByZip2Tax = "<font size=3>Sales Tax<b> " & TempRate & "%</b></font><br />( for zip code " & Zipx & " )<br /><br />Details:<br />" & StateX & "<br />" & CountyX & "<br />" & CityX
    LookUpTaxByZip2Tax = LookUpTaxByZip2Tax & "<br /><br />Tax Rate Information retrived from Zip2Tax.com<br />To<b> Copy</b> or<b> Save</b> (as text), use buttons below"

    Call MessageUser("Close")
    Call Dialog.RichBox(LookUpTaxByZip2Tax, vbCustom + vbInformation, "Tax Lookup - Zip:" & Zipx, , , 0, 1, 1, 1, "OK")
End Function
Function CleanTrim(ByVal s As String, Optional ConvertNonBreakingSpace As Boolean = True) As String
  Dim x As Long, CodesToClean As Variant
  CodesToClean = "0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, _
                       21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 127, 129, 141, 143, 144, 157)
  If ConvertNonBreakingSpace Then s = Replace(s, Chr(160), " ")
  For x = LBound(CodesToClean) To UBound(CodesToClean)
    If InStr(s, Chr(CodesToClean(x))) Then s = Replace(s, Chr(CodesToClean(x)), "")
  Next
  CleanTrim = Trim(s)
End Function
