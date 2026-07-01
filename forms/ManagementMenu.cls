Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim SQLz As String
    Dim XLS, Headerx, SQLy As String

Private Sub AccountingLoanReport_Click()
    DoCmd.OpenForm "LoanReportParameters"
End Sub

Private Sub BDexPerdiem_Click()
    SQLy = "SELECT * FROM PerDiem_Days_Rate_BDEX"
    MessageUser "Open", "BDEX PD Report", "Excel will open soon..."
    Headerx = "PerDiem By BDEX - Account - " & Date
    XLS = "|A,BC|F,SE|F,BC|G,$E|H,$E|I,$E|"
   ' Countx = DCount("UnitID", "PerDiem_Days_Rate_BDEX")
    SendReport2Excel SQLy, "Perdiem Unit", Headerx, True, XLS, 0
    MessageUser "Close"
End Sub

Private Sub BillingReport_Click()
  If Me.StartDateRange & "test" = "test" Or Me.EndDateRange & "test" = "test" Then
        MsgBox "Sorry... You must select at lease a Date range", vbCritical, "No Dates"
        Exit Sub
    End If
    Call ExportBilling(Me.StartDateRange, Me.EndDateRange, "Bank", Nz(Me.SelectClient, 0))
End Sub

Private Sub BLCD_NoDelivery__Click()
    Headerx = ""
    SQLz = "Select ClientShortName as Client, Schedule, UnitGroup As [Group], BankShortName as Assignee"
    SQLz = SQLz & ", AssetID, UnitNum, VIN, [Status], UnitDescShortVer as UDesc, BLCD"
    SQLz = SQLz & ", ActualDeliveryDate as Delivered, AcceptedDate as Accepted, InServiceDate as InServiced, FundedDate"
    SQLz = SQLz & " From vw_SixKeys"
    SQLz = SQLz & " Where Not BLCD Is Null"
    SQLz = SQLz & " AND ActualDeliveryDate IS NULL"
    SQLz = SQLz & " AND NOT FundedDate IS NULL"
    If Nz(Me.ClientList, 0) <> 0 Then
       SQLz = SQLz & " AND ClientId=" & Me.ClientList.Column(0)
       Headerx = Me.ClientList.Column(1) & " - "
    End If
    SQLz = SQLz & " Order By ClientShortName"
    Headerx = Headerx & "Groups With BLCD but No Delivery - As of:" & Date
    XLS = "|E,TX|E,BC|H,BC|J,BC|K,cColor:3:215.150.35|"
    SendReport2Excel SQLz, "Fleet List", Headerx, True, XLS, 0, True
End Sub

Private Sub AssigneeLocationsExport_Click()
    SQLz = "SELECT "
    SQLz = SQLz & "BankShortName As Assignee"   'a
    SQLz = SQLz & ",SchBnkClNO & ""-"" & SgrpBnkNo AS Contract " ' b
    SQLz = SQLz & ",CLientShortName As Client"  'c
    SQLz = SQLz & ",Schedule" 'd
    SQLz = SQLz & ",AssetID As [FA AssetID]"      'e
    SQLz = SQLz & ",UnitNum" 'f
    SQLz = SQLz & ",VIN" 'g
    SQLz = SQLz & ",CurrRent" 'h
    SQLz = SQLz & ",S.LocationSn As Domicile" 'i
    SQLz = SQLz & ",Nz(ActualDeliveryDate,H.LocationChangeDate) AS [Date Domiciled] " ' j
    SQLz = SQLz & ",L.LCounty as County" 'k
    SQLz = SQLz & ",S.LState as State" ' l
    SQLz = SQLz & ",S.LZip as ZIP" ' m
    SQLz = SQLz & ",S.[Status] " ' n
    SQLz = SQLz & " FROM (((vw_SixKeys As S "
    SQLz = SQLz & " INNER JOIN Locations AS L ON S.LocationID = L.LocationID) "
    SQLz = SQLz & " INNER JOIN SchGrp ON S.SGrpID = SchGrp.SGrpID) "
    SQLz = SQLz & " INNER JOIN Schedule ON S.SchID = Schedule.SchID)"
    SQLz = SQLz & " LEFT OUTER JOIN (SELECT AssetId as hAssetID, NewLocationId, LocationChangeDate  "
    SQLz = SQLz & " FROM LocationChangeHistory) As H ON S.AssetID = H.hAssetID and S.LocationID = H.NewLocationId"
    SQLz = SQLz & " WHERE ([Status] = 'A' OR [Status] = 'TR')"
    If Me.SelectClient <> "" Then SQLz = SQLz & " AND S.ClientGroupID = " & Me.SelectClient.Column(0) & " "
    If Nz(Me.SelectBank, 0) > 0 Then SQLz = SQLz & " AND S.BankId = " & Me.SelectBank.Column(0) & " " Else SQLz = SQLz & " AND S.BankID = 4198 "
    SQLz = SQLz & "Order By BankShortName, CLientShortName "
    Headerx = "Units at Location, By Assignee: " & Date
    XLS = "|A,DD3|A,W12|E,TX|E,cColor:3:black|E,cBorder:3:Med|E,TC|D,TC|H,$E|J,TC|J,W11|L,BC|N,TC|"
    SendReport2Excel SQLz, "Units-Location", Headerx, True, XLS, 0, True

End Sub

Private Sub btn180ExpireReport_Click()
    Call Make180ExpireReport
End Sub

Private Sub EXITAMGExcel_Click()
        Dim ClientGrpi As Long
        Dim ClientGroupName As String
        Dim BankX As String
        If Me.SelectClient.Column(0) & "test" <> "test" Then
        ClientGrpi = Me.SelectClient.Column(0)
        ClientGroupName = Me.SelectClient.Column(1)
        End If
        If Me.SelectBank.Column(1) & "test" <> "test" Then BankX = Me.SelectBank.Column(1)

        Call MakeExitAMG(ClientGrpi, ClientGroupName, BankX)
End Sub


Private Sub RvRealizationRpt_Click()
   Call MakeRVRealizedReport(Nz(Me.SelectClient, 0), Nz(Me.SelectBank, 0))
End Sub

Private Sub HistoricFleetList_Click()
    If Me.StartDateRange & "test" = "test" Or Me.EndDateRange & "test" = "test" Or Me.SelectClient & "test" = "test" Then
        MsgBox "Sorry... You must select a Client and Date range", vbCritical, "No Clien, No Dates"
        Exit Sub
    End If

    Headerx = Me.SelectClient.Column(1) & " Fleet List From During - " & StartDateRange & " ~ " & EndDateRange
    XLS = "|D,BC|D,TX|G,TC|L,TC|O,TC|P,BC|Q,TC|R,TC|S,BC|T,SintE|U,SE|V,SintE|W,SintR|X,SR|"

    SQLz = "SELECT Client, Schedule, UnitGroup, AssetID, Unitnum, VIN, MYear, Make, Model, Type, SubType "
    SQLz = SQLz & ", InServiceDate, Locationsn, Division, OffLeaseActual, cStatus "
    SQLz = SQLz & ", MIN(LogActivityDate) AS FirstReport,MAX(LogActivityDate) AS LastReport, Count(AssetID) as Reports "
    SQLz = SQLz & ", SUM(h.Miles) AS Miles, SUM(h.Fuel) AS Fuel, SUM(h.FuelMoving) AS FuelMoving "
    SQLz = SQLz & ", MIN(h.ODOStart) AS OdoStart, MAX(h.ODOEnd) AS OdoEnd, UnitDesc "
    SQLz = SQLz & "FROM vw_FleetList_OBC as h "
    SQLz = SQLz & "WHERE ClientGroupID = " & Me.SelectClient & " "
    SQLz = SQLz & "AND InServiceDate <= #" & EndDateRange & "# "
    SQLz = SQLz & "AND (OffLeaseActual is null or OffLeaseActual >= #" & EndDateRange & "#) "
    SQLz = SQLz & "AND LogActivityDate >= #" & StartDateRange & "# "
    SQLz = SQLz & "AND LogActivityDate <= #" & EndDateRange & "# "
    SQLz = SQLz & "AND OBCDataStatus < 3 "
    SQLz = SQLz & "GROUP BY Client, Schedule, UnitGroup, AssetID, Unitnum, VIN, MYear, Make, Model, Type, SubType"
    SQLz = SQLz & ",InServiceDate, Locationsn, Division, OffLeaseActual, cStatus, UnitDesc, ClientGroupId "
    SQLz = SQLz & "ORDER BY Make, MYear DESC"
    MessageUser "Open", "Getting Historic Fleet List", "Excel will open soon..."
    SendReport2Excel SQLz, "FleetList", Headerx, True, XLS, 0
    MessageUser "Close"
End Sub

Private Sub Imported_LPAFS_Click()
    If FAUserAccessGroup Like "*TMT*" Or FAUserAccessGroup Like "*ITT*" Then
        DoCmd.close acForm, "ATF_Structure"
        DoCmd.OpenForm "ATF_Structure"
    End If

End Sub

Private Sub MilesReport_Click()
    SQLz = "Select Client" '- Col:A
    SQLz = SQLz & " ,Sch" '- Col:B
    SQLz = SQLz & " ,Group" '- Col:C
    SQLz = SQLz & " ,AssetID" '- Col:D
    SQLz = SQLz & " ,UnitNum" '- Col:E
    SQLz = SQLz & " ,MYear" '- Col:F
    SQLz = SQLz & " ,Make" '- Col:G
    SQLz = SQLz & " ,Model" '- Col:H
    SQLz = SQLz & " ,SubType" '- Col:I
    SQLz = SQLz & " ,Type" '- Col:J
    SQLz = SQLz & " ,Status" '- Col:K
    SQLz = SQLz & " ,Axle" '- Col:L
    SQLz = SQLz & " ,VIN" '- Col:M
    SQLz = SQLz & " ,AgeBLCD" '- Col:N
    SQLz = SQLz & " ,AgeInServ" '- Col:O
    SQLz = SQLz & " ,ODODate" '- Col:P
    SQLz = SQLz & " ,AgeODO" '- Col:Q
    SQLz = SQLz & " ,InServODOGap" '- Col:R
    SQLz = SQLz & " ,LeaseStatus" '- Col:S
    SQLz = SQLz & " ,LeaseTermPri" '- Col:T
    SQLz = SQLz & " ,LeaseTermExtd" '- Col:U
    SQLz = SQLz & " ,AllowedPerYear" '- Col:V
    SQLz = SQLz & " ,AllowedMiles" '- Col:W
    SQLz = SQLz & " ,ActualMiles" '- Col:X
    SQLz = SQLz & " ,AllowedMilesLeft" '- Col:Y
    SQLz = SQLz & " ,AllowedMilesUsedPct" '- Col:Z
    SQLz = SQLz & " ,LeaseMthsLeft" '- Col:AA
    SQLz = SQLz & " ,LeaseTermUsedPct" '- Col:AB
    SQLz = SQLz & " ,ActualMilesPerMonth" '- Col:AC
    SQLz = SQLz & " ,ActualMilesSource" '- Col:AD
    SQLz = SQLz & " ,ProjMilesToday" '- Col:AE
    SQLz = SQLz & " ,ProjMilesEOL" '- Col:AF
    SQLz = SQLz & " ,MileCharge" '- Col:AG
    SQLz = SQLz & " ,OffLeaseProj" '- Col:AH
    SQLz = SQLz & " ,OffLeaseActual" '- Col:AI
    SQLz = SQLz & " ,SoldDate" '- Col:AJ
    SQLz = SQLz & " ,AgeSold" '- Col:AK
    SQLz = SQLz & " ,SoldPrice" '- Col:AL
    SQLz = SQLz & " ,UnitRV" '- Col:AM
    SQLz = SQLz & " ,ClientGroupId" '- Col:AN
    SQLz = SQLz & " FROM vw_ExitMilesReport "
    Headerx = "Actual Miles VS Miles Allowed"
    If Me.SelectClient & "test" <> "test" Then
        SQLz = SQLz & " Where ClientGroupID= " & Me.SelectClient & " "
        Headerx = Headerx & " Client: " & Me.SelectClient.Column(1)
    End If
    SQLz = SQLz & "Order By Client, AssetID "
    Headerx = Headerx & " as of: " & Date
    XLS = "|A,Wrap|"
    XLS = XLS & "A,DD3|A,TC|A,cLabel:End=A:3:Client|A,cColor:3:0.0.0|"
    XLS = XLS & "B,TC|B,cLabel:End=B:3:Sch|B,cColor:3:0.0.0|"
    XLS = XLS & "C,TC|C,cLabel:End=C:3:Group|C,cColor:3:0.0.0|"
    XLS = XLS & "D,TX|D,TC|D,cLabel:End=D:3:AssetID|D,cColor:3:0.0.0|"
    XLS = XLS & "E,TC|E,cLabel:End=E:3:Unit Num|E,cColor:3:0.0.0|"
    XLS = XLS & "F,TC|F,cLabel:End=F:3:MYear|F,cColor:3:0.0.0|"
    XLS = XLS & "G,TC|G,cLabel:End=G:3:Make|G,cColor:3:0.0.0|"
    XLS = XLS & "H,TC|H,cLabel:End=H:3:Model|H,cColor:3:0.0.0|"
    XLS = XLS & "I,TC|I,cLabel:End=I:3:SubType|I,cColor:3:0.0.0|"
    XLS = XLS & "J,TC|J,cLabel:End=J:3:Type|J,cColor:3:0.0.0|"
    XLS = XLS & "K,BC|K,cLabel:End=K:3:Status|K,cColor:3:0.0.0|"
    XLS = XLS & "L,TC|L,cLabel:End=L:3:Axle|L,cColor:3:0.0.0|"
    XLS = XLS & "M,TC|M,cLabel:End=M:3:VIN|M,cColor:3:0.0.0|"
    XLS = XLS & "N,Sint|N,TC|N,cLabel:End=N:3:AgeBLCD|N,cColor:3:0.0.0|"
    XLS = XLS & "O,Sint|O,TC|O,cLabel:End=O:3:Age In Service|O,cColor:3:0.0.0|"
    XLS = XLS & "P,TC|P,W11|P,cLabel:End=P:3:Odometer Date|P,cColor:3:68.114.196|"
    XLS = XLS & "Q,Sint|Q,BC|Q,W11|Q,cLabel:End=Q:3:Age Odometer|Q,cColor:3:47.117.181|"
    XLS = XLS & "R,Sint|R,TC|R,W11|R,cLabel:End=R:3:InServ ODO Gap (mths)|R,cColor:3:55.86.35|"
    XLS = XLS & "S,TC|S,W11|S,cLabel:End=S:3:Lease Status|S,cColor:3:0.0.0|"
    XLS = XLS & "T,Sint|T,TC|T,W11|T,cLabel:End=T:3:Primary Lease  Term|T,cColor:3:0.0.0|"
    XLS = XLS & "U,Sint|U,TC|U,W11|U,cLabel:End=U:3:Lease Term Extd|U,cColor:3:89.89.89|"
    XLS = XLS & "V,Sint|V,TC|V,W11|V,cLabel:End=V:3:Allowed Miles Per Year|"
    XLS = XLS & "W,Sint|W,TC|W,W11|W,cLabel:End=W:3:Allowed Miles To Date|"
    XLS = XLS & "X,Sint|X,0|X,W11|X,cLabel:End=X:3:Actual Miles Driven Odometer|X,cColor:3:68.114.196|"
    XLS = XLS & "Y,Sint|Y,0|Y,W11|Y,cLabel:End=Y:3:Allowed Miles Left|"
    XLS = XLS & "Z,P1|Z,BC|Z,W11|Z,cLabel:End=Z:3:Allowed Miles Used Pct|"
    XLS = XLS & "AA,Sint|AA,TC|AA,W11|AA,cLabel:End=AA:3:Lease Mths Left|AA,cColor:3:89.89.89|"
    XLS = XLS & "AB,P1|AB,BC|AB,W11|AB,cLabel:End=AB:3:Lease Term Used Pct|AB,cColor:3:89.89.89|"
    XLS = XLS & "AC,Sint|AC,TC|AC,W11|AC,cLabel:End=AC:3:Miles Driven Per Month (avg)|"
    XLS = XLS & "AD,TC|AD,W11|AD,cLabel:End=AD:3:Odometer Source|AD,cColor:3:68.114.196|"
    XLS = XLS & "AE,Sint|AE,TC|AE,W11|AE,cLabel:End=AE:3:Miles To Date (Active = projected)|AE,cColor:3:102.0.51|"
    XLS = XLS & "AF,Sint|AF,TC|AF,W11|AF,cLabel:End=AF:3:Miles at EOL (Active = projected) |AF,cColor:3:102.0.51|"
    XLS = XLS & "AG,$R|AG,0|AG,cLabel:End=AG:3:Per Mile Charge|"
    XLS = XLS & "AH,TC|AH,W11|AH,cLabel:End=AH:3:Off Lease Proj|"
    XLS = XLS & "AI,TC|AI,W11|AI,cLabel:End=AI:3:Off Lease Actual|AI,cColor:3:204.51.0|"
    XLS = XLS & "AJ,TC|AJ,W11|AJ,cLabel:End=AJ:3:Sold Date|AJ,cColor:3:0.0.0|"
    XLS = XLS & "AK,P1|AK,TC|AK,cLabel:End=AK:3:Age Sold|AK,cColor:3:0.0.0|"
    XLS = XLS & "AL,$R|AL,TC|AL,cLabel:End=AL:3:Sold Price|AL,cColor:3:0.0.0|"
    XLS = XLS & "AM,$R|AM,TC|AM,cLabel:End=AM:3:Unit RV|AM,cColor:3:0.0.0|"
    XLS = XLS & "AN,HH|AN,cLabel:End=AN:3:ClientGroupId|"


    MessageUser "Open", "Getting Miles Report", "Excel will open soon..."
    SendReport2Excel SQLz, "MilesDriven", Headerx, True, XLS, 0
    MessageUser "Close"
End Sub

Private Sub PayProceedsManager_Click()
    DoCmd.close acForm, "PayProceedsManager"
    DoCmd.OpenForm "PayProceedsManager"
End Sub
Private Sub OverMiles_MPY25_Click()
    SQLz = "SELECT Client"
    SQLz = SQLz & ", Sch"
    SQLz = SQLz & ", Grp"
    SQLz = SQLz & ", AssetID"
    SQLz = SQLz & ", [Unit Num]"
    SQLz = SQLz & ", [Unit Desc]"
    SQLz = SQLz & ", BLCD"
    SQLz = SQLz & ", PriTerm"
    SQLz = SQLz & ", [Lease Age]"
    SQLz = SQLz & ", [Mths Left]"
    SQLz = SQLz & ", [ODO Read Date]"
    SQLz = SQLz & ", [ODO Source]"
    SQLz = SQLz & ", Odometer"
    SQLz = SQLz & ", [MPY Allowed]"
    SQLz = SQLz & ", [MPY Used]"
    SQLz = SQLz & ", [MPY Over]"
    SQLz = SQLz & ", [Pct Ovr]"
    SQLz = SQLz & ", Domicile"
    SQLz = SQLz & " FROM  vw_OverMiles_ExceedsMPY"
    SQLz = SQLz & " WHERE [Mths Left] >0 "
    SQLz = SQLz & " AND   [Pct Ovr] > 0.24"
     Headerx = "Units Excceding Allowed MPY By 25% Or More"
    If Me.SelectClient & "test" <> "test" Then
        SQLz = SQLz & " AND   ClientGroupId = " & Me.SelectClient & " "
        Headerx = "Client: " & Me.SelectClient.Column(1) & " : Units Excceding Allowed MPY By 25% Or More"
    End If
    SQLz = SQLz & " ORDER BY Client, [Sub Type], [Pct Ovr] desc "

    XLS = ""
    XLS = XLS & "|A,DD3|A,W15|B,TC|C,TC|D,BC|D,TX|E,TC|G,TC|H,TC|I,TC|I,W07|J,TC|J,W07"
    XLS = XLS & "|K,TC|K,W10|L,TC|L,W08|M,Sint|N,Sint|N,W10|O,Sint|O,W10|P,Sint|P,W10|Q,%|Q,TC|"

     MessageUser "Open", "Getting Over Alloerd MPY Report", "Excel will open soon..."
    SendReport2Excel SQLz, "Excceds MPY", Headerx, True, XLS, 0
    MessageUser "Close"

End Sub
Private Sub ReLeasedUnits_Click()
    Headerx = "Re-Leased Units"
    SQLz = "Select Client" '- Col:
    SQLz = SQLz & " ,Schedule" '- Col:B
    SQLz = SQLz & " ,UnitGroup" '- Col:C
    SQLz = SQLz & " ,BLCD" '- Col:D
    SQLz = SQLz & " ,OffLease" '- Col:E
    SQLz = SQLz & " ,BuyOutDate" '- Col:F
    SQLz = SQLz & " ,BuyOutCost" '- Col:G
    SQLz = SQLz & " ,Status" '- Col:H
    SQLz = SQLz & " ,Unitnum" '- Col:I
    SQLz = SQLz & " ,Assignee" '- Col:J
    SQLz = SQLz & " ,AssetID" '- Col:K
    SQLz = SQLz & " ,VIN" '- Col:L
    SQLz = SQLz & " ,UnitDesc" '- Col:M
    SQLz = SQLz & " ,NewAssetID" '- Col:N
    SQLz = SQLz & " ,NewBank" '- Col:O
    SQLz = SQLz & " ,NewUnitNum" '- Col:P
    SQLz = SQLz & " ,NewStatus" '- Col:Q
    SQLz = SQLz & " ,NewVendorCost" '- Col:R
    SQLz = SQLz & " ,NewLeaseExp" '- Col:S
    SQLz = SQLz & " ,NewBLCD" '- Col:T
    SQLz = SQLz & " ,NewGroup" '- Col:U
    SQLz = SQLz & " ,NewSchedule" '- Col:V
    SQLz = SQLz & " ,NewClient" '- Col:W
    SQLz = SQLz & " ,AssetID_FATitleLog" '- Col:X
    SQLz = SQLz & " ,TitleToTrust" '- Col:Y

    SQLz = SQLz & " From vw_ReLeasedUnits "
    If Me.SelectClient & "test" <> "test" Then
        SQLz = SQLz & " Where ClientGroupID= " & Me.SelectClient & " "
        Headerx = Headerx & "- Client: " & Me.SelectClient.Column(1)
    End If
    SQLz = SQLz & "Order By Client, AssetID "
    MessageUser "Open", "Getting Report", "Excel will open soon..."

    XLS = ""
    XLS = XLS & "A,DD3|A,TC|A,cLabel:End=A:3:Client|"
    XLS = XLS & "B,TC|B,cLabel:End=B:3:Sch|"
    XLS = XLS & "C,TC|C,cLabel:End=C:3:Group|"
    XLS = XLS & "D,TC|D,cLabel:End=D:3:BLCD|"
    XLS = XLS & "E,TC|E,W11|E,cLabel:End=E:3:OffLease (Actual)|"
    XLS = XLS & "F,TC|F,cLabel:End=F:3:Buy Out (Date)|F,cColor:3:47.117.181|F,W11|"
    XLS = XLS & "G,$E|G,cLabel:End=G:3:Buy Out Actual|"
    XLS = XLS & "H,BC|H,W07|H,cLabel:End=H:3:Status|"
    XLS = XLS & "I,TC|I,W09|I,cLabel:End=I:3:Unitnum|"
    XLS = XLS & "J,TC|J,W09|J,cLabel:End=J:3:Assignee|"
    XLS = XLS & "K,TX|K,BC|K,W09|K,cLabel:End=K:3:AssetID|K,cColor:3:31.78.120|"
    XLS = XLS & "L,TC|L,cLabel:End=L:3:VIN|L,cColor:3:0.0.0|"
    XLS = XLS & "M,TC|M,cLabel:End=M:3:UnitDesc|M,cColor:3:0.0.0|"
    XLS = XLS & "N,BC|N,W09|N,cLabel:End=N:3:New AssetID|N,cColor:3:55.86.35|"
    XLS = XLS & "O,TC|O,W09|O,cLabel:End=O:3:New Bank|O,cColor:3:84.130.53|"
    XLS = XLS & "P,TC|P,W09|P,cLabel:End=P:3:New UnitNum|P,cColor:3:84.130.53|"
    XLS = XLS & "Q,BC|Q,W07|Q,cLabel:End=Q:3:New Status|Q,cColor:3:84.130.53|"
    XLS = XLS & "R,$E|R,W13|R,cLabel:End=R:3:New Vendor $|R,cColor:3:55.86.35|"
    XLS = XLS & "S,BC|S,W11|S,cLabel:End=S:3:New Lease Expire|S,cColor:3:84.130.53|"
    XLS = XLS & "T,TC|T,W11|T,cLabel:End=T:3:New BLCD|T,cColor:3:84.130.53|"
    XLS = XLS & "U,TC|U,W11|U,cLabel:End=U:3:New Grp|U,cColor:3:84.130.53|"
    XLS = XLS & "V,TC|V,W11|V,cLabel:End=V:3:New Sch|V,cColor:3:84.130.53|"
    XLS = XLS & "W,TC|W,W11|W,cLabel:End=W:3:New Client|W,cColor:3:84.130.53|"
    XLS = XLS & "X,TC|X,0|X,cLabel:End=X:3:FATitleLog|X,cColor:3:117.113.113|"
    XLS = XLS & "Y,TC|Y,0|Y,cLabel:End=Y:3:TitleToTrust|Y,cColor:3:117.113.113|"


    XLS = XLS & "A,DD3|A,Wrap|"
    SendReport2Excel SQLz, "Re-LeaseUnits", Headerx, True, XLS, 0
    MessageUser "Close"
End Sub

Private Sub SelectBank_GotFocus()
    If Nz(Me.SelectClient, 0) = 0 Then
        Me.SelectBank.RowSource = "SELECT Banks.BankId, Banks.BankShortName FROM Banks ORDER BY Banks.BankShortName;"
    Else
        Me.SelectBank.RowSource = "SELECT AssigneeBankID, Assignee From vw_ExitReport0 WHERE ClientGroupID=" & Me.SelectClient & " Group By AssigneeBankID, Assignee ORDER BY Assignee;"
    End If
End Sub

Private Sub SyndByState_Click()

SQLz = "SELECT Client "     ' Col A
SQLz = SQLz & ",Schedule"   ' Col B
SQLz = SQLz & ",SyndState"  ' Col C
SQLz = SQLz & ",SyndYear"   ' Col D
SQLz = SQLz & ",OEC"        ' Col E
SQLz = SQLz & ",SyndFee"    ' Col F
SQLz = SQLz & ",PerDiem"    ' Col G
SQLz = SQLz & ",Units"      ' Col H
SQLz = SQLz & ",SyndDate"   ' Col I
SQLz = SQLz & ",Rent"       ' Col J
SQLz = SQLz & ",SPDRent as [Surr Rent]"    ' Col K
'SQLz = SQLz & ",ClientGroupId"
'SQLz = SQLz & ",ClientID"
'SQLz = SQLz & ",SchID "
SQLz = SQLz & " FROM vw_SyndByState As S "
SQLz = SQLz & " WHERE 1 = 1"
If Nz(Me.SelectClient, 0) > 0 Then SQLz = SQLz & " AND S.ClientGroupID = " & Me.SelectClient & " "
If Nz(StartDateRange, "") <> "" And Nz(EndDateRange, "") <> "" Then
   SQLz = SQLz & " AND SyndDate >= #" & Me.StartDateRange & "# "
   SQLz = SQLz & " AND SyndDate <= #" & Me.EndDateRange & "# "
End If
SQLz = SQLz & " Order By Client, SyndDate, Schedule"
XLS = "|C,BC|D,BC|E,$E|F,$E|G,$E|H,BC|J,$E|K,$E|"

SendReport2Excel SQLz, "Synd By State", "Synd By State", True, XLS, 0

End Sub

Private Sub BDVReport_Click()
    If Me.StartDateRange & "test" = "test" Or Me.EndDateRange & "test" = "test" Then
        MsgBox "Sorry... You must set a date range", vbCritical, "No Dates"
        Exit Sub
    End If
    MessageUser "Open", "Getting BDVP Report", "Excel will open soon..."
    Call AddReportLog("Business Devlopement")
    SQLy = "SELECT SentYear, Client, FAID,DealBEDEX as BDEX, DealYear, DealUnitQty, DealGrandTotal, Units, OEC, UnitDiff, OECDiff, OutcomeDate From BDVOP_Excel2 " & _
           "WHERE OutcomeDate Between #" & Me.StartDateRange & "# And #" & Me.EndDateRange & "#;"
    Headerx = "Business Development Volume - " & StartDateRange & " ~ " & EndDateRange
    XLS = "|A,BC|D,BC|E,BC|F,NE|G,$E|H,NE|I,$E|J,NE|K,$E|"
    Countx = DCount("DealYear", "BDVOP_Excel2")

    SendReport2Excel SQLy, "BDVP Report", Headerx, True, XLS, 0
    MessageUser "Close"

End Sub

Private Sub ClearAll_Click()
    Me.StartDateRange = Null
    Me.EndDateRange = Null
    Me.SelectBank = Null
    Me.SelectClient = Null
    Me.UnitSelection = 1
End Sub

Private Sub ClientListByYear_Click()
    MessageUser "Open", "Getting Clients By Year", "Excel will open, just a sec..."
    SQLz = "SELECT"
    SQLz = SQLz & " S.ClientShortName AS Client, S.BankShortName AS Bank, D.Dyear, S.MLNo, S.Schedule & ' | ' & S.UnitGroup AS SchGrp, S.AssetsInGroup as Units, D.FAID"
    SQLz = SQLz & " FROM  vw_SixKeys AS S LEFT OUTER JOIN"
    SQLz = SQLz & "      (SELECT FAID, DATEPART('YYYY', Nz(dealOutcomeDate, dealEnteredDate)) AS Dyear"
    SQLz = SQLz & "       FROM  Deals) AS D ON D.FAID = S.FAID"
    SQLz = SQLz & " WHERE (S.MLOrig = 'FA')"
    SQLz = SQLz & " GROUP BY S.BankShortName, S.ClientShortName, S.MLNo, S.Schedule, S.UnitGroup, S.AssetsInGroup, S.SGrpID, D.Dyear, D.FAID"
    SQLz = SQLz & " ORDER BY D.Dyear, ClientShortName,BankShortName"

    'SQLz = "SELECT BankShortName As Bank, ClientShortName as Client, MLNO, Schedule & ' | ' & UnitGroup As SchGrp, AssetsInGroup " & _
     "FROM vw_SixKeys WHERE Mlorig='FA'" & _
     "Group By BankShortName , ClientShortName , MLNO, Schedule, UnitGroup, AssetsInGroup"
    SendReport2Excel SQLz, "CByYear", "Client List By Year Report ", True, "|A,BC|C,BC|F,BC|", 0, True
    MessageUser "Close"
End Sub

Private Sub CloseMe_Click()
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub AssigneeActivity_Click()
    If Me.StartDateRange & "test" = "test" Then
    MsgBox "You Must Select a Start/End Date to run this report", vbCritical, "Missing Parameters"
    Exit Sub
    End If
    MessageUser "Open", "Getting Assignee Activity", "Excel will open, just a sec..."
    SQLz = "SELECT * FROM vw_AssigneeActivity " & _
           "WHERE Funded Between #" & StartDateRange & "# AND #" & Nz(EndDateRange, "12/31/2050") & "# Order By Assignee, CompanyName, Funded;"
    SendReport2Excel SQLz, "Funded", "Assignee Activity By Funded Date Report", True, "|A,BL|F,SE|H,$E|", 0, True
    MessageUser "Close"
End Sub

Private Sub Ctl_TFRDashboard_Click()
    DoCmd.close acForm, "TFRDashboard"
    DoCmd.OpenForm "TFRDashboard"
    Call AddReportLog("TFR Dashboard")
End Sub

Private Sub TFRReportSavePDF_Click()
    Call MakeTFRReport("Report", Nz(Me.StartDateRange, ""), Nz(EndDateRange, ""), Nz(Me.SelectClient, 0), Nz(Me.SelectBank, 0), True)
End Sub

Private Sub ExitOptionsDashboard_Click()
    DoCmd.OpenForm "ExitOptionsManager"

End Sub

Private Sub Loan2Excel_Click()
If Me.StartDateRange & "test" = "test" Then
MsgBox "You Must Select a Start Date before you can run a report", vbCritical, "Missing Parameters"
End If
   SQLz = "Select Client, MLNo, Schedule, UnitGroup As [Group], Assignee, AssetID, VIN, UnitStatus as Status, Cost as OEM " & _
          ",Vendor_wParts as [Vendor+], LoanDate, LoanBank, LoanAmount, UnitPayOff as PayOff " & _
          "From InterimLoanq " & _
          "WHERE (PayoffDate is null) OR (LoanDate >=#" & Me.[StartDateRange] & " # AND PayoffDate>=#" & Me.[StartDateRange] & "#)"
  XLS = "|A,BL|F,BL|H,BC|I,$E|J,$E|L,BL|M,$E|"
Call SendReport2Excel(SQLz, "Loan Report", "Loan Report (Interim) As Of " & Me.StartDateRange, True, XLS, 0, True)
End Sub

Private Sub NewClientByYear_Click()
SQLz = "Select * From vw_ClientDeals Order By DYear Desc, Client;"
  XLS = "|A,BC|D,BC|G,HH|"
  Call SendReport2Excel(SQLz, "ClientDeals", "New Client By Year", True, XLS, 0, True)

End Sub

Private Sub OpenPerposalOutcomes_Click()
    DoCmd.OpenReport "ProposalsOutcomes"
End Sub

Private Sub RVSyndExport_Click()
    If Nz(Me.SelectClient, 0) > 0 Then
        Dim ClientGrpi As Long
        Dim Clienti As Long
        Dim ClientGroupName As String
        ClientGrpi = Nz(Me.SelectClient.Column(0), 0)
        ClientGroupName = Me.SelectClient.Column(1)
        If Nz(Me.SelectClient.Column(2), 0) = 27 Then
            ClientGrpi = Me.SelectClient.Column(2)
            Clienti = 0
            ClientGroupName = "Maines Group"
        End If
        Call MakeRVSyndReport(Clienti, ClientGrpi, ClientGroupName)
    Else
        Call MakeRVSyndReport
    End If
End Sub

Private Sub TFR_TMT_Click()
    If Me.StartDateRange & "test" = "test" Or Me.EndDateRange & "test" = "test" Then
        MsgBox "Sorry... You must set a date range", vbCritical, "No Dates"
        Exit Sub
    End If
    Call MessageUser("Open", "Getting TMT's TFR", "Just a sec...")
    SQLz = " SELECT "
    SQLz = SQLz & "AccountName"
    SQLz = SQLz & ", [Schedule] + ' | ' + [UnitGroup] AS SchGrp"
    SQLz = SQLz & ", FundDate"
    SQLz = SQLz & ", Bank"
    SQLz = SQLz & ", LeaseTerm"
    SQLz = SQLz & ", Avg(MilesAllowed) AS MilesPYear"
    SQLz = SQLz & ", UnitCount As Units"
    SQLz = SQLz & ", MYear"
    SQLz = SQLz & ", Make"
    SQLz = SQLz & ", Model" 'j
    SQLz = SQLz & ", Type"
    SQLz = SQLz & ", SubType"
    SQLz = SQLz & ", Axle"
    SQLz = SQLz & ", OECPerUnit"
    SQLz = SQLz & ", VendorCostUnit" ' o
    SQLz = SQLz & ", PDRentActual"
    SQLz = SQLz & ", Round([PDRentActual]/[TotalVendorCost],4) AS NetPDPct"
    SQLz = SQLz & ", RV"
    SQLz = SQLz & ", RVPct"
    SQLz = SQLz & ", SWAPType" 't
    SQLz = SQLz & ", Yield"  'u
    SQLz = SQLz & ", IndexRate" ' v
    SQLz = SQLz & ", Spread"    ' w
    SQLz = SQLz & ", SpreadIRR" 'x
    SQLz = SQLz & ", YieldIRR"  'y
    SQLz = SQLz & ", Bonus " 'z
    SQLz = SQLz & "From TFR_2 "
    SQLz = SQLz & "GROUP BY AccountName, [Schedule] + ' | ' + [UnitGroup],FundDate,Bank,LeaseTerm,UnitCount"
    SQLz = SQLz & ",MYear,Make,Model,Type,SubType,Axle,OECPerUnit,VendorCostUnit,PDRentActual"
    SQLz = SQLz & ",Round([PDRentActual]/[TotalVendorCost],4),RV,RVPct,SWAPType,Yield,IndexRate,Spread,SpreadIRR,YieldIRR,Bonus "
    SQLz = SQLz & "HAVING FundDate Between #" & StartDateRange & "# AND #" & EndDateRange & "#;"
    Headerx = "TMT's TFR Report: " & StartDateRange & "~" & EndDateRange & ""
    XLS = "|A,DD|A,cLabel:End=A:3:Account" & vbNewLine & "Name|"
    XLS = XLS & "B,TC|"
    XLS = XLS & "D,BL|"
    XLS = XLS & "E,TC|E,cLabel:End=E:3:Lease" & vbNewLine & "Term|"
    XLS = XLS & "F,Sint|F,cLabel:End=F:3:Miles" & vbNewLine & "Per Yr|"
    XLS = XLS & "G,BC|"
    XLS = XLS & "N,$R|N,cLabel:End=N:3:OEC" & vbNewLine & "Per Unit|"
    XLS = XLS & "O,$R|O,cLabel:End=O:3:Vendor" & vbNewLine & "Per Unit|"
    XLS = XLS & "P,$R|P,cLabel:End=P:3:PDRent" & vbNewLine & "Per Unit|"
    XLS = XLS & "Q,P2|"
    XLS = XLS & "R,$R|"
    XLS = XLS & "S,P2|"
    XLS = XLS & "T,BC|"
    XLS = XLS & "U,P2|"
    XLS = XLS & "V,P2|"
    XLS = XLS & "W,P2|"
    XLS = XLS & "X,P2|"
    XLS = XLS & "Y,P2|"
    XLS = XLS & "Z,P1|"
    SendReport2Excel SQLz, "TFR", Headerx, True, XLS, 0, False
    Call MessageUser("Close")
End Sub

Private Sub TFRShortReport_Click()
  Call MakeTFRShort(Nz(Me.StartDateRange, ""), Nz(EndDateRange, ""), Nz(Me.SelectClient.Column(0), 0), Nz(Me.SelectBank, 0))
End Sub
Private Sub TMTTFRReport_Click()

  Call MakeTMTTFR(Nz(Me.StartDateRange, ""), Nz(EndDateRange, ""), Nz(Me.SelectClient.Column(0), 0), Nz(Me.SelectBank, 0), Me.FundStatus)

End Sub

Private Sub TFRSummaryReport_Click()
'DoCmd.Close acReport, "TFRCompareSummary"
'DoCmd.OpenReport "TFRCompareSummary", acViewPreview
    Call MakeTFRReport("Report", "TFRCompareSummary", Nz(Me.StartDateRange, ""), Nz(EndDateRange, ""), Nz(Me.SelectClient, 0), Nz(Me.SelectBank, 0))
End Sub

Private Sub ExchangeReport_Click()

    Dim UnitsX, Active, Built, OffLease, RMKT, Sold, Track As Integer
    Dim RST As DAO.Recordset
    On Error GoTo ClientFleetList_Click_Error

    'UnitCount = DCount("[AssetID]", "vw_ClientFleetList", "[ClientId]=" & Me.clientID)
    Call MessageUser("Open", "Creating Report", " ...This will take just a moment...")

    If Nz(Me.ClientList, 0) = 0 Then
        SQLz = "SELECT * From FleetUnitCountByStatus_All WHERE "
    Else
        SQLz = "SELECT * From FleetUnitCountByStatus WHERE ClientID=" & Me.ClientList & " AND "
    End If

    Set RST = CurrentDb.OpenRecordset(SQLz & " Status = 'A'")
    If Not RST.EOF Then Active = RST!UnitCount
    RST.close
    Set RST = CurrentDb.OpenRecordset(SQLz & " Status = 'B'")
    If Not RST.EOF Then Built = RST!UnitCount
    RST.close
    Set RST = CurrentDb.OpenRecordset(SQLz & "Status like '*R*'")
    If Not RST.EOF Then RMKT = RST!UnitCount
    RST.close

    Set RST = CurrentDb.OpenRecordset(SQLz & "Status = 'C'")
    If Not RST.EOF Then OffLease = RST!UnitCount
    RST.close
    Set RST = CurrentDb.OpenRecordset(SQLz & "Status = 'O'")
    If Not RST.EOF Then OffLease = Nz(OffLease, 0) + RST!UnitCount
    RST.close
    Call MessageUser("Update", "Collecting Unit counts...", " ...Just another moment or two...")

    Set RST = CurrentDb.OpenRecordset(SQLz & " Status like '*S*'")
    If Not RST.EOF Then Sold = RST!UnitCount
    RST.close

    Set RST = CurrentDb.OpenRecordset(SQLz & " Status = 'T'")
    If Not RST.EOF Then Track = RST!UnitCount
    RST.close

    If Nz(Me.ClientList, 0) = 0 Then
        SQLz = "SELECT Sum(FleetUnitCountByStatus.UnitCount) AS Units FROM FleetUnitCountByStatus;"
    Else
        SQLz = "SELECT clientID, Sum(UnitCount) AS Units  From FleetUnitCountByStatus GROUP BY clientID Having ClientID =" & Me.ClientList & ";"
    End If

    If Nz(Me.ClientList, 0) = 0 Then
        SQLz = "SELECT * From Client_FleetList_Exchange Order By Client"
        Headerx = "All Clients' Fleet List -"
    Else
        SQLz = "Select * From Client_FleetList_Exchange Where ClientId=" & Me.ClientList & ";"
        Headerx = Me.ClientList.Column(1) & " Fleet List -"
    End If
UnitsX = Nz(Active, 0) + Nz(RMKT, 0) + Nz(Track, 0)
    If UnitsX <= 0 Then
        MsgBox "Sorry, But " & Me.ClientList.Column(1) & " dosen't have a Fleet List", vbCritical, "No Units"
        DoCmd.close acForm, "MessagePopUp"
        Exit Sub
    End If
    Call MessageUser("Update", "Done doing Unit counts...", " ...Just a Bit More...")
    If Active > 0 Then Headerx = Headerx & " Active: " & Active
    If Built > 0 Then Headerx = Headerx & " Being Built: " & Built
    If OffLease > 0 Then Headerx = Headerx & " Off Lease: " & OffLease
    If RMKT > 0 Then Headerx = Headerx & " Rmkt: " & RMKT
    If Sold > 0 Then Headerx = Headerx & " Sold: " & Sold
    If Track > 0 Then Headerx = Headerx & " Tracked: " & Track
    Headerx = Headerx & " As of:" & Date
    XLS = "|B,BC|H,CX|K,BC|O,BL|Q,BC|R,$E|S,$E|T,$E|U,$E|V,$E|W,P5C|X,$E|Z,BC|AB,$E|AD,$E|AE,HH|AF,HH|W,W14|"
    SendReport2Excel SQLz, "Fleet List", Headerx, True, CustomScript, 0, True
    Call MessageUser("Close")

    On Error GoTo 0
    Exit Sub

ClientFleetList_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ClientFleetList_Click of VBA Document Form_Mgmt Report"
End Sub
Private Sub Exit2Excel_Click()
    Call RmktGainLossFix(True)
    Call MakeExitExcelReport(0, Nz(Me.SelectClient.Column(0), 0), Nz(Me.SelectClient.Column(1), ""), , , Me.UnitSelection, False, CLng(Nz(Me.SelectBank, 0)))
End Sub

Private Sub ExitReportMDA_Click()
    Call RmktGainLossFix(True)
    Call MakeExitMDAReport(0, Nz(Me.SelectClient.Column(0), 0), Nz(Me.SelectClient.Column(1), ""), , , 2, CLng(Nz(Me.SelectBank, 0)))
End Sub

Private Sub ExitShort_Click()
    Call RmktGainLossFix(True)
    Call MakeExitExcelReport(0, Nz(Me.SelectClient.Column(0), 0), Nz(Me.SelectClient.Column(1), ""), , , Me.UnitSelection, True, CLng(Nz(Me.SelectBank, 0)))
End Sub
Private Sub BDEXExitReport_Click()
    Dim MainText, Subtext As String
    Dim LastDay As Date
    Dim RST As DAO.Recordset
    Dim Rst2 As DAO.Recordset
    Dim ii, AllRecords, RecCountx, TabX, NLoop As Integer
    Dim IsLastRpt, ShowXL As Boolean
    ' On Error GoTo Fixme
    MessageUser "Open", "Excel - BDEX Exit Report", "Just a moment or two..."
    LastDay = CDate("12/31/" & DatePart("yyyy", Date) + 1)
    SQLz = "SELECT SP.SalespersonID, SP.SalespersonName AS BDEXName, BDEX, SP.SalespersonEmail AS BDEXEmail, Count(AssetID) as Units "
    SQLz = SQLz & "FROM vw_SixKeys As Six INNER JOIN Salesperson As SP ON Six.BDEX = SP.SPInitials "
    SQLz = SQLz & "WHERE (Status = 'A' Or Status = 'TR' Or Status = 'P')"
    'SQLz = SQLz & "WHERE Six.Status = 'A' "
    If Me.ClientList <> "" Then
        SQLz = SQLz & "AND ClientGroupID=" & Me.ClientList.Column(2) & " "
    Else
        SQLz = SQLz & "AND LeaseExp <= #" & LastDay & "# "
    End If
    SQLz = SQLz & "GROUP BY SP.SalespersonID, SP.SalespersonName, BDEX, SP.SalespersonEmail "
    SQLz = SQLz & "Order By BDEX desc"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    RST.MoveLast
    RST.MoveFirst
    NLoop = RST.RecordCount
    'Get Total Records
    For ii = 1 To NLoop
        AllRecords = AllRecords + RST!UNITS
        RST.MoveNext
    Next ii

    'Reset for Next Loop
    ii = 0
    RST.MoveFirst

    'Get Exit Report Loop
    For ii = 0 To NLoop
        SQLz = "SELECT "
        SQLz = SQLz & "Client "
        SQLz = SQLz & ",BDEX"
        SQLz = SQLz & ",Sch"
        SQLz = SQLz & ",[Group]"
        SQLz = SQLz & ",Garage As [Location]"
        SQLz = SQLz & ",State As LState "
        SQLz = SQLz & ",InService"
        SQLz = SQLz & ",Unitnum"
        SQLz = SQLz & ",[Status]"
        SQLz = SQLz & ",UnitVin as VIN"
        SQLz = SQLz & ",Axle"
        SQLz = SQLz & ",BLCD"
        SQLz = SQLz & ",LeaseTerm as Term"
        SQLz = SQLz & ",LeaseExp as ExpiryDate"
        SQLz = SQLz & ",ExchangeITMonth as ExchangetITTerm"
        SQLz = SQLz & ",ExchangeITDate as [1stExchangeIT]"
        SQLz = SQLz & ",ExpireNotice as NoticeDate"
        SQLz = SQLz & ",MthsInServ"
        SQLz = SQLz & ",UnitYr As MYear"
        SQLz = SQLz & ",UnitMake As Make"
        SQLz = SQLz & ",UnitModel As Model"
        SQLz = SQLz & ",UnitSubType As SubType"
        SQLz = SQLz & ",UnitType As Type "
        SQLz = SQLz & "From vw_ExitReport0 as E "
        SQLz = SQLz & "WHERE MLOrig = 'FA' "
        SQLz = SQLz & "AND (Status = 'A' Or Status = 'TR' Or Status = 'P')"
        ' SQLz = SQLz & "AND LeaseExp <= #" & LastDay & "# "
        If Me.ClientList <> "" Then
            SQLz = SQLz & "AND ClientGroupID=" & Me.ClientList.Column(2) & " "
        Else
            SQLz = SQLz & "AND LeaseExp <= #" & LastDay & "# "
        End If
        '    If Me.ClientList <> "" Then SQLz = SQLz & "AND ClientGroupID=" & Me.ClientList.Column(2) & " "
        If ii > 0 Then SQLz = SQLz & "AND BDEX='" & RST!BDEX & "' "
        SQLz = SQLz & " Order By Client;"


        If ii < 1 Then
            Headerx = "EXIT Report - Master - " & Date
            MainText = "Master"
            IsLastRpt = False
            ShowXL = False
            Subtext = "There will be Several Tabs..."
            TabX = -1
            RecCountx = AllRecords
        Else
            MainText = RST!BDEXname
            Subtext = (NLoop - ii) + 1 & " tabs to go..."
            IsLastRpt = False
            ShowXL = False
            TabX = ii
            RecCountx = RST!UNITS
        End If

        If ii = NLoop Then
            IsLastRpt = True
            ShowXL = True
            Subtext = "Excel Will Open Soon"
            TabX = 6
        End If
        '  RecCountx = rst2.RecordCount
        Headerx = "EXIT Report - " & MainText & " - " & Date
        MessageUser "Update", "Making " & MainText & " Tab", Subtext
        XLS = "|B,BC|F,BC|I,BC|I,CX|M,TC|N,TR|O,TC|P,TR|R,TC|"
        XLS = XLS & "N,VColor:" & RecCountx + 3 & ":185.211.238|"
        XLS = XLS & "P,VColor:" & RecCountx + 3 & ":185.211.238|"
        If RecCountx > 1 Then SendReport2Excel SQLz, MainText, Headerx, True, XLS, 0, ShowXL, TabX, IsLastRpt
MoveNextRec:
        If ii > 0 Then RST.MoveNext
    Next ii
FixMe:
    MessageUser "Close"
    On Error Resume Next
    RST.close
    Rst2.close
    Set RST = Nothing
    Set Rst2 = Nothing
    On Error GoTo 0
End Sub
Private Sub ExitOptions_Click()
    DoCmd.OpenForm "ExitOptions"
End Sub


Private Sub ExitReportDetail_Click()
    DoCmd.close acReport, "ExitReport"
    DoCmd.OpenReport "ExitReport", acViewReport
End Sub

Private Sub Expiry_Click()
    Dim SQLBdex, SQLExit, SQLEol, TabText, Subtext, TitleX, Pre_Title As String
    Dim RST As DAO.Recordset
    Dim ii, ee, ReportNum, RecCountx As Integer
    Dim IsLastRpt, ShowXL As Boolean

    On Error Resume Next
    DoCmd.SetWarnings False
    DoCmd.DeleteObject Table, "BuyOut_Temp"
    On Error GoTo 0
    Call MessageUser("Open", "Collecting Data", "This will take a moment")
    SQLz = "Select * INTO BuyOut_Temp From vw_BuyOutCalcUnit Where BDEX <> '' "
    SQLz = SQLz & SQLEol & " Order By BDEX, Client;"
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    TitleX = ""

    If IsDate(Me.StartDateRange) Then TitleX = TitleX & Me.StartDateRange & " "
    If IsDate(Me.StartDateRange) And IsDate(Me.EndDateRange) Then TitleX = TitleX & "~ "
    If IsDate(Me.EndDateRange) Then TitleX = TitleX & Me.EndDateRange & " "

    If IsDate(Me.StartDateRange) Then SQLEol = "AND ExpiryDate >=#" & Me.StartDateRange & "# "
    If IsDate(Me.EndDateRange) Then SQLEol = SQLEol & "AND ExpiryDate <=#" & Me.EndDateRange & "# "

    If IsDate(Me.StartDateRange) Then SQLExit = "AND ExchangeITDate >=#" & Me.StartDateRange & "# "
    If IsDate(Me.EndDateRange) Then SQLExit = SQLExit & "AND ExchangeITDate <=#" & Me.EndDateRange & "# "

    SQLEol = SQLEol & SqlClient & SqlBank
    Call MessageUser("Update", "Setting Dates and BDEX", "This will be quick...")
MakeBDEX:
    SQLBdex = "SELECT BDEX, SalespersonName " & _
              "FROM BuyOut_Temp INNER JOIN Salesperson ON BuyOut_Temp.BDEX = Salesperson.SPInitials " & _
              "WHERE BDEX <> 'HS' AND BDEX <> 'JM' AND BDEX <> 'PF' "
    SQLBdex = SQLBdex & SQLEol & " GROUP BY BDEX, SalespersonName Order By BDEX Desc"
    Set RST = CurrentDb.OpenRecordset(SQLBdex, dbOpenSnapshot)
    RST.MoveLast
    RST.MoveFirst
    RecCountx = RST.RecordCount + 2
    For ii = 1 To RecCountx
        ' If ii <= RST.RecordCount Then SQLz = "Select * From BuyOut_Temp WHERE BDEX='" & RST!BDEX & "'"
        For ee = 1 To 2
            XLS = "|C,BC|D,RX|F,TC|H,TC|J,$E|K,$E|L,$E|M,$E|N,$E|O,$E|"
            SQLz = "Select * From BuyOut_Temp WHERE BDEX"
            If RecCountx - ii = 1 Then
                SQLz = SQLz & "= 'HS' or BDEX ='JM' or BDEX = 'PF' "
                If ee = 1 Then TabText = "HS_EOL" Else TabText = "HS_Exchange"
                If ee = 1 Then Pre_Title = "EOL Report - House" Else Pre_Title = "EXchangeIT Report - House"
            End If
            If RecCountx = ii Then
                SQLz = SQLz & "<>''"
                If ee = 1 Then TabText = "ALL_EOL" Else TabText = "ALL_Exchange"
                If ee = 1 Then Pre_Title = "EOL Report - ALL" Else Pre_Title = "EXchangeIT Report - ALL"
                If ee = 2 Then ShowXL = True Else ShowXL = False
                If ee = 2 Then IsLastRpt = True Else IsLastRpt = False
            End If
            If ii <= RST.RecordCount Then
                XLS = XLS & "C,ZZ|"
                SQLz = SQLz & "='" & RST!BDEX & "'"
                If ee = 1 Then TabText = RST!BDEX & "_EOL" Else TabText = RST!BDEX & "_Exchange"
                If ee = 1 Then Pre_Title = "EOL Report - " & RST!SalespersonName Else Pre_Title = "EXchangeIT Report - " & RST!SalespersonName
                IsLastRpt = False
                ShowXL = False
            End If
            If ee = 1 Then SQLz = SQLz & SQLEol
            If ee = 2 Then SQLz = SQLz & SQLExit
            ReportNum = ReportNum + 1
            Call MessageUser("Update", "Making Tab " & TabText, "This is Report " & ReportNum & " of " & RecCountx * 2)
            SendReport2Excel SQLz & " Order By Client", TabText, Pre_Title & " - " & TitleX, True, XLS, 0, ShowXL, ReportNum, IsLastRpt
        Next ee
        If ii < RST.RecordCount Then RST.MoveNext
    Next ii
    RST.close
    Set RST = Nothing
    DoCmd.SetWarnings False
    DoCmd.DeleteObject Table, "BuyOut_Temp"
    DoCmd.SetWarnings True
    Call MessageUser("Close")
End Sub

Private Sub Form_Load()
    If FAUserAccess > 14 Or FaUserName() = "kzeigler" Or FaUserName() = "lredmond" Then
        'Me.AccountingLoanReport.Visible = True
        Me.SyndExportReport.Visible = True
        Me.ClientList.Visible = True
        Me.ExchangeReport.Visible = True
    Else
       ' Me.AccountingLoanReport.Visible = False
        Me.SyndExportReport.Visible = False
        Me.ClientList.Visible = False
        Me.ExchangeReport.Visible = False
    End If
    If FAUserAccess > 74 Then Me.TFRReport.Visible = True
    Me.YTD0.caption = DatePart("yyyy", Date)
    Me.YTD1.caption = DatePart("yyyy", Date) - 1
    Me.YTD2.caption = DatePart("yyyy", Date) - 2
    Me.StartDateRange = ""
    Me.EndDateRange = ""
    Call RmktGainLossFix(True)
    ' Call YTD0_Click
End Sub

Private Sub LeaseProfileReport_Click()
    Dim XLS As String
    XLS = "|A,DD|B,TX|F,BC|J,$E|K,$R|L,P3|M,$E|N,P5|P,BC|Q,TC|R,$R|T,BC|U,$R|V,$E|W,$E|Y,ZZ|"

    Call MessageUser("Open", "Lease Summary", "Working on the Off Lease Tab")
    SQLz = "Select * From vw_LeasedUnitsProfile WHERE nz([Unit Status],'') = 'A' or nz([Unit Status],'') = 'B' or nz([Unit Status],'') = 'TR' or nz([Unit Status],'') = 'P' Order By Customer"
    Call SendReport2Excel(SQLz, "Active", "Fleet Advantage - Off Lease Summary " & Date, True, XLS, 0, False, 1, False)

    Call MessageUser("Update", "Lease Summary", "Working on the Active Tab...")
    SQLz = "Select * From vw_LeasedUnitsProfile WHERE nz([Unit Status],'') = 'C' or nz([Unit Status],'') = 'R' or nz([Unit Status],'') = 'Ti' or nz([Unit Status],'') = 'PS' or nz([Unit Status],'') = 'S' or nz([Unit Status],'') = 'D' Order By Customer"
    Call SendReport2Excel(SQLz, "Off Lease", "Fleet Advantage - Active Lease Summary " & Date, True, XLS, 0, True, 2, True)
    Call AddReportLog("Lease Summary Report")
    Call MessageUser("Close")
End Sub

Private Sub LeasesByYear_Click()
    DoCmd.close acReport, "LeaseReportByYear"
    Call AddReportLog("LeaseReportByYear")
    DoCmd.OpenReport "LeaseReportByYear", acViewReport
End Sub

Private Sub ATFOutStanding_Schedules_Click()
    DoCmd.close acReport, "OutstandingATFSchedule_KPI_Report"
    DoCmd.OpenReport "OutstandingATFSchedule_KPI_Report", acViewReport, , , acWindowNormal
End Sub
Private Sub ATFCompleted_Schedules_Click()
    DoCmd.close acReport, "CompletedATFSchedule_KPI_Report"
    DoCmd.OpenReport "CompletedATFSchedule_KPI_Report", acViewReport, , , acWindowNormal
End Sub

Private Sub ROnsiteBtn_Click()
    Dim stDocName As String

    stDocName = "ProfileRank"

    DoCmd.SetWarnings False

    DoCmd.RunMacro stDocName
    DoCmd.SetWarnings True
End Sub

Private Sub SelectClient_Click()
    Me.ClientGroup = Me.SelectClient.Column(1)
End Sub

Private Sub SyndExportReport_Click()
   ' Dim Countx As Integer
    On Error GoTo SyndExportReport_Click_Error

    MessageUser "Open", "Setting up Funded Tab", "this will take a moment"
    XLS = ""
    XLS = "|A,DD|A,BC|B,BC|E,BC|I,LX|J,TC|M,$E|O,$E|P,P2R|Q,$E|R,P2R|S,$E|T,P2R|U,$E|V,P2R|W,$E|X,P2R|Y,$E|Z,P2R|" & _
               "AA,$E|AB,P2R|AC,$E|AD,P2R|AE,$E|AF,P2R|AG,$E|AH,P2R|AJ,$E|AK,P2R|AL,$E|AM,P2R|AN,$E|AO,P2R|AP,$E|AQ,P2R|AR,$E|AS,P2R|AT,$E|AU,P2P|AV,P2R|AW,P2R|" & _
               "H,W13|J,W08|N,W12|O,W14|P,W08|Q,W14|R,W08|S,W14|T,W10|U,W14|V,W08|W,W14|X,W10|Y,W14|Z,W08|" & _
               "AA,W14|AB,W08|AC,W14|AD,W08|AE,W14|AF,W08|AG,W14|AH,W08|AI,W14|AJ,W14|AK,W10|AL,W14|AM,W10|AN,W14|AO,W08|AP,W14|AQ,W08|AR,W14|AS,W08|AT,W15|AU,W10|AW,W10|AX,W10|AY,W10|BB,$E|BC,$E|"

    SQLz = "Select * From Syndication_Validate_Portfolio Where [Fund Status]= 'F' Order By [CompanyName];"
    'Countx = DCount("AssetID", "Syndication_Validate_Portfolio", "[Fund Status]='F'")

    MessageUser "Update", "Portfolio Report", "Exporting Funded Portfolio Tab. working...", True
    SendReport2Excel SQLz, "Funded", "Funded Portfolio Report | " & Date, True, XLS, 0, False, 1, False

    MessageUser "Update", "Portfolio Report", "Exporting Hold Tab. working...", True
    SQLz = "Select * From Syndication_Validate_Portfolio Where [Fund Status]= 'H' Order By [CompanyName];"
    'Countx = DCount("AssetID", "Syndication_Validate_Portfolio", "[Fund Status]='H'")
    SendReport2Excel SQLz, "Hold", "Funded Portfolio Report | " & Date, True, XLS, 0, False, 2, False

    MessageUser "Update", "Portfolio Report", "Exporting Inventory Tab. working...", True
    SQLz = "Select * From Syndication_Validate_Portfolio Where [Fund Status]= 'I' Order By [CompanyName];"
    'Countx = DCount("AssetID", "Syndication_Validate_Portfolio", "[Fund Status]='I'")
    SendReport2Excel SQLz, "Inventory", "Funded Portfolio Report | " & Date, True, XLS, 0, False, 3, False

    MessageUser "Update", "Portfolio Report", "Exporting Forecast Portfolio. almost done...", True
    SQLz = "Select * From Syndication_Validate_Portfolio_Pending Where [Fund Status]='P' Order By [CompanyName];"
    'Countx = DCount("AssetID", "Syndication_Validate_Portfolio_Pending", "[Fund Status]='P'")
    SendReport2Excel SQLz, "Forecast", "Forecast Portfolio Report | " & Date, True, XLS, 0, True, 4, True

    MessageUser "Close", "", ""

    On Error GoTo 0
    Exit Sub

SyndExportReport_Click_Error:
    MessageUser "Close", "", ""
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SyndExportReport_Click of VBA Document Form_ManagementMenu"
End Sub

Private Sub TFRPreView_Click()
    Call MakeTFRReport("Print", Nz(Me.StartDateRange, ""), Nz(EndDateRange, ""), Nz(Me.SelectClient, 0), Nz(Me.SelectBank, 0))
End Sub
Private Sub TFRReport_Click()
    Call MakeTFRReport("Report", "TFRCompare", Nz(Me.StartDateRange, ""), Nz(EndDateRange, ""), Nz(Me.SelectClient, 0), Nz(Me.SelectBank, 0))
End Sub

Private Sub TFROffSiteExcel_Click()
    Call TFRExcel(Nz(Me.StartDateRange, ""), Nz(EndDateRange, ""), Nz(Me.SelectClient.Column(0), 0), Nz(Me.SelectBank, 0))
End Sub

Private Sub CustodianDashboard_Click()
    DoCmd.OpenForm "TitleToCustodianDashboard"
End Sub




Private Sub VendorDelvReport_Click()

    MessageUser "Open", "Getting Vendor Report", "...Excel will open in a moment..."
    SQLz = "SELECT * "
    SQLz = SQLz & "FROM vw_VendorFundReport "
    If Me.StartDateRange & "test" <> "test" Then SQLz = SQLz & "WHERE [VendorPaid] >= #" & Me.StartDateRange & "# And [VendorPaid] <= #" & Me.EndDateRange & "# "
    SQLz = SQLz & " ORDER BY Assignee, Vendor, Make;"
    XLS = "|D,CE|E,$E|F,$E|G,$E|L,BC|"
    SendReport2Excel SQLz, "Vendors", "Vendor Volume: " & Me.StartDateRange & " ~ " & Me.EndDateRange & "", True, XLS, 0, 0
    MessageUser "Close"
End Sub

Private Sub YTD0_Click()
    Call SetDateRange("StartDateRange", "EndDateRange", 0)
End Sub
Private Sub YTD1_Click()
    Call SetDateRange("StartDateRange", "EndDateRange", 1)
End Sub
Private Sub YTD2_Click()
    Call SetDateRange("StartDateRange", "EndDateRange", 2)
End Sub
Public Sub SetButtonYears()
    Dim x As Integer
    Dim bName As String
    For x = 0 To 2
        bName = "YTD" & x
        Me(bName).caption = Right(DatePart("yyyy", Date) - x, 4)
    Next x
End Sub
Private Sub SetDateRange(StartField As String, EndField As String, YearBack As Integer)
    Me(StartField) = "7/1/" & DatePart("yyyy", Date) - (YearBack + 1)
    'If YearBack = 0 Then Me(EndField) = Date Else
    Me(EndField) = "6/30/" & DatePart("yyyy", Date) - YearBack
End Sub
