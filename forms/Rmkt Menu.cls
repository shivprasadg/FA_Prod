Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz, SQLx, XLS, Header As String



Private Sub BuyersOrdersMassMod_Click()
    DoCmd.close acForm, "BuyersOrderMassMod"
    DoCmd.OpenForm "BuyersOrderMassMod"
End Sub

Private Sub ClearAll_Click()
    Me.InvoiceNum = ""
    Me.StartDate = ""
    Me.EndDate = ""
    Me.Vendor = ""
    Me.Bank = ""
End Sub

Private Sub CloseMe_Click()
    DoCmd.OpenForm "FAMainMenu"
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub Customer_Delivery_Report_Click()
    DoCmd.OpenReport "CustomerDeliveryReportDetailed", acViewPreview
End Sub
Public Sub SetSQLx()
    SQLx = "SELECT VendorName, InvoiceNumber, InvoiceDate, ExpenseDate, Expense, RemarketExpense AS ExpType, Comment "
    SQLx = SQLx & ", ClientShortname AS Client,BankShortName AS Assignee, vw_SixKeys.AssetID, UnitNum, VIN, Status "
    SQLx = SQLx & ", UnitDescShortVer, LocationSn, RMKTExpenseId "
    SQLx = SQLx & "FROM ((vw_SixKeys "
    SQLx = SQLx & "INNER JOIN RemarketExpense ON vw_SixKeys.AssetID = RemarketExpense.AssetId) "
    SQLx = SQLx & "INNER JOIN Vendors ON RemarketExpense.VendorId = Vendors.VendorID) "
    SQLx = SQLx & "INNER JOIN RemarketExpenseType ON RemarketExpense.ExpenseTypeId = RemarketExpenseType.RemarketExpenseTypeId "
    XLS = "|A,TL|B,BL|E,$E|H,TX|J,BC|K,TC|M,BC|"
End Sub



Private Sub Exit2Excel_Click()
    If Nz(Me.StartDate, "1/1/1900") <> "1/1/1900" And Nz(Me.EndDate, "1/1/1900") <> "1/1/1900" Then

        Call MakeExitExcelReport(, , , Me.StartDate, Me.EndDate, 4)
    Else
        Call MakeExitExcelReport(, , , , , 4)
    End If
End Sub

Private Sub InterimStatusReport_Click()
    If FAUserFullName = "Ysis Jackson" Or FAUserAccessLevel = 99 Then
        Call RMKTInterimStatusReport
    Else
        MsgBox "Sorry you do not have access", vbInformation, "Freeze!"
    End If
End Sub

Private Sub MakeSUFEmail_Click()
    If Emailto = 1 Then fSufReport Else fSufReport (True)
End Sub

Private Sub MonthlySoldReport_Click()
 If Not IsDate(Me.StartDate) Or Not IsDate(Me.EndDate) Then
        MsgBox "You must have BOTH a Start and End Date Set", vbCritical, "No Dates, No Service"
        Exit Sub
 End If
    Call Remarketing_SOLDReport_Sales("UnitStatus in ('S','D') AND InvoicePaid >=#" & Me.StartDate & "# AND InvoicePaid<=#" & Me.EndDate & "# ")
End Sub

Private Sub BTF_BuyersOrder_Click()
  'ACCT-ALL-ITT-UAT-TMT-SYND-PUR
    Dim grpAccess As String
    grpAccess = FAUserAccessGroup
    If grpAccess Like "*TMT*" _
    Or grpAccess Like "*ITT*" _
    Or grpAccess Like "*ACCT*" _
    Or FAUserName = "ytj" _
    Or FAUserName = "fmaloney" _
    Or FAUserName = "tvanahn" Then
       DoCmd.close acForm, "BTF_BuyersOrder"
       DoCmd.OpenForm "BTF_BuyersOrder"
    Else
        Call MsgBox("Sorry You Do Not Have Accesss", vbCritical, "Slow down partner")
    End If
End Sub
Private Sub OpenRmktExpenseManager_Click()
    DoCmd.OpenForm "RmktExpenseManager"
End Sub

Private Sub OpenUnitMoverManager_Click()
  DoCmd.close acForm, "UnitMoveOrder_Maker"
  DoCmd.OpenForm "UnitMoveOrder_Maker"
End Sub

Private Sub Outstanding_SUFs_Click()
    Call Remarketing_OutStanding_SUF
End Sub

Private Sub BankInventoryReport_Click()
   ' If Me.Bank & "test" = "test" Then
   '     MsgBox "You must Select a Bank", vbCritical, "No Bank, I LOVE YOU"
   '     Exit Sub
   ' End If
   Call RMKT_BankInventoryReport(Nz(Me.Bank.Column(0), 0), Nz(Me.Bank.Column(1), ""))
End Sub

Private Sub RMKTBankReport_Click()
    If Me.Bank & "test" = "test" Then
        MsgBox "You must Select a Bank", vbCritical, "No Bank, No Service"
        Exit Sub
    End If
    Header = "Expenses By Assignee: " & Me.Bank.Column(1)
    MessageUser "Open", "Creating RMKT Invoice Report", "...Getting Data..."
    Call SetSQLx
    SQLz = SQLx
    SQLz = SQLz & "WHERE BankID=" & Me.Bank & " "
    If IsDate(Me.StartDate) And IsDate(Me.EndDate) Then
        SQLz = SQLz & "AND  ExpenseDate >=#" & Me.StartDate & "# AND  ExpenseDate <=#" & Me.EndDate & "# "
        Header = Header & " and By Date Range: " & Me.StartDate & " ~ " & Me.EndDate
    End If
    SQLz = SQLz & "ORDER BY InvoiceNumber, ExpenseDate, ClientShortname, vw_SixKeys.AssetID"
    SendReport2Excel SQLz, "Vendor", Header, True, XLS, 0
    MessageUser "Close"
End Sub
Private Sub LookUpRMKTInvoice_Click()
    If InvoiceNum & "test" = "test" Then
        MsgBox "You must have an Invoice Number Entered", vbCritical, "No Invoice, No Service"
        Exit Sub
    End If
    If DCount("RMKTExpenseID", "RemarketExpense", "InvoiceNumber='" & Me.InvoiceNum & "'") = 0 Then
        MsgBox "There are NO matching Invoices that have the" & vbNewLine & "Invoice Number: " & Me.InvoiceNum, vbCritical, "No Match"
        Exit Sub
    End If
    Header = "Look Up Invoice Num: " & Me.InvoiceNum
    MessageUser "Open", "Creating RMKT Invoice Report", "...Getting Data..."
    Call SetSQLx
    SQLz = SQLx & "WHERE InvoiceNumber='" & Me.InvoiceNum & "' "
    SQLz = SQLz & "ORDER BY InvoiceNumber, ExpenseDate, VendorName, vw_SixKeys.AssetID "
    SendReport2Excel SQLz, "Invoice", Header, True, XLS, 0
    MessageUser "Close"
End Sub

Private Sub RMKTVendorReport_Click()
    If Vendor & "test" = "test" Then
        MsgBox "You must Select a Vendor", vbCritical, "No Vendor, No Service"
        Exit Sub
    End If
    Header = "Expenses By Vendor: " & Me.Vendor.Column(1)
    MessageUser "Open", "Creating RMKT Invoice Report", "...Getting Data..."
    Call SetSQLx
    SQLz = SQLx
    SQLz = SQLz & "WHERE Vendors.VendorID=" & Me.Vendor & " "
    If IsDate(Me.StartDate) And IsDate(Me.EndDate) Then
        SQLz = SQLz & "AND  ExpenseDate >=#" & Me.StartDate & "# AND  ExpenseDate <=#" & Me.EndDate & "# "
        Header = Header & " and By Date Range: " & Me.StartDate & " ~ " & Me.EndDate
    End If
    SQLz = SQLz & "ORDER BY InvoiceNumber, ExpenseDate, ClientShortname, vw_SixKeys.AssetID"
    SendReport2Excel SQLz, "Vendor", Header, True, XLS, 0
    MessageUser "Close"
End Sub
Private Sub RMKTExpenseDate_Report_Click()
    If Not IsDate(Me.StartDate) Or Not IsDate(Me.EndDate) Then
        MsgBox "You must have BOTH a Start and End Date Set", vbCritical, "No Dates, No Service"
        Exit Sub
    End If
    Header = "Expenses By Date: " & Me.StartDate & " ~ " & Me.EndDate
    MessageUser "Open", "Creating RMKT Invoice Date Report", "...Getting Data..."
    Call SetSQLx
    SQLz = SQLx
    SQLz = SQLz & "WHERE ExpenseDate >=#" & Me.StartDate & "# AND  ExpenseDate <=#" & Me.EndDate & "# "
    SQLz = SQLz & "ORDER BY ExpenseDate, VendorName, ClientShortname, vw_SixKeys.AssetID"
    SendReport2Excel SQLz, "Expense Dates", Header, True, XLS, 0
    MessageUser "Close"
End Sub


Private Sub MakeBankPayOffEmail_Click()
    Dim Committ As String
    Dim Mstring As String
    Mstring = "You're about send the 'Bank Pay-Off Invoice Request' Emails to "
    If Emailto = 2 Then
        Mstring = Mstring & vbNewLine & " Just " & FAUserFullName & vbNewLine & "Do you agree?"
    Else
        Mstring = Mstring & "Ysis." & vbNewLine & "Do you agree?"
    End If
    Committ = MsgBox(Mstring, vbYesNo, "Think! " & TragetX & "?")
    If Committ = vbNo Then Exit Sub
    If Emailto = 2 Then Call SendRequestPayOffEmail(1) Else Call SendRequestPayOffEmail(0)
End Sub

Private Sub InventoryManagerView_Click()
    DoCmd.OpenForm "RemarketingReportMakerMGR"
End Sub



Private Sub Form_Load()
    Call ClearAll_Click
   '' Call FixParkedUnits(True)
    Me.EndDate = Date + 5
    Me.StartDate = DateAdd("d", -30, Date)


    Me.OverRideCbo = "No OverRide"
    If FAUserAccess < 70 Then
        '   Me.GetDekra1.Visible = False
        '   Me.startDate.Visible = False
        '   Me.Enddate.Visible = False
    Else
       'Me.GetDekra1.Visible = True
        Me.StartDate.Visible = True
          Me.EndDate.Visible = True
    End If
    If FAUserAccess < 80 Then
        Me.OverRideCbo.Visible = False
        Me.SPDashboard.Visible = False
         Me.RangeLabel.Visible = False
            Exit2Excel.Visible = False
    Else
        Me.OverRideCbo.Visible = True
        Me.SPDashboard.Visible = True
         Me.RangeLabel.Visible = True
            Exit2Excel.Visible = True
    End If
    If FAUserAccessGroup Like "*Acct*" Then Me.OpenDashBoardAcct.Visible = True Else Me.OpenDashBoardAcct.Visible = False
    If FAUserAccess < 15 Then
        Me.OpenRmktExpenseManager.Visible = False
           Me.OpendashBoardInvMgr.Visible = False
            Me.OpenDashBoardAdmin.Visible = False
               Me.OpenRMKTMassMod.Visible = False

    Else
        Me.OpenRmktExpenseManager.Visible = True
           Me.OpendashBoardInvMgr.Visible = True
            Me.OpenDashBoardAdmin.Visible = True
               Me.OpenRMKTMassMod.Visible = True

    End If
End Sub

Private Sub GetDekra1_Click()
    '  Call DownloadDekraDataExcel
    '   Me.Form.Refresh
End Sub

Private Sub StartDate_AfterUpdate()
   Me.EndDate = DateAdd("m", 1, Me.StartDate) - 1
End Sub

Private Sub SUFReturned_Click()
    If Not IsDate(Me.StartDate) Or Not IsDate(Me.EndDate) Then
        MsgBox "You must have BOTH a Start and End Date Set", vbCritical, "No Dates, No Service"
        Exit Sub
    End If
    SQLz = "Select Client"                           'A
    SQLz = SQLz & ", Sch"                            'B
    SQLz = SQLz & ", [Group] as Grp"                 'C
    SQLz = SQLz & ", Assignee"                       'D
    SQLz = SQLz & ", AssetID"                        'E
    SQLz = SQLz & ", UnitNum"                        'F
    SQLz = SQLz & ", UnitVIn"                        'G
    SQLz = SQLz & ", ParkedDate as [Parked Date]"    'H
    SQLz = SQLz & ", SUFBack"                        'I
    SQLz = SQLz & ", ActualOffLease as OffLeased"    'J
    SQLz = SQLz & ", ParkedMiles as [Miles Parked]"  'K
    SQLz = SQLz & ", HoursParked as [Hours Parked]"  'L
    SQLz = SQLz & ", InspectionDate as [Insp Date]"  'M
    SQLz = SQLz & ", InspectionPassed as [Insp Pass]"     'N
    SQLz = SQLz & ", UnitStatus As Status"                'O
    SQLz = SQLz & ", ODORead AS [ODO Read]"               'P
    SQLz = SQLz & ", ODOSource AS [ODO Source]"           'Q
    SQLz = SQLz & ", 0 As [Miles Delta]"                  'R
    SQLz = SQLz & ", MilesAllowedAll AS [Miles Allowed]"  'S
    SQLz = SQLz & ", OverMiles AS [Miles Over-Undr]"      'T
    SQLz = SQLz & ", OverMilesCost AS [Over Mile Charge]" 'U
    SQLz = SQLz & " From vw_ExitReport0 as E"
    SQLz = SQLz & " Where Mlorig='FA' AND ParkedDate is not null"
    SQLz = SQLz & " AND  SUFBack >=#" & Me.StartDate & "# AND  SUFBack <=#" & Me.EndDate & "# "
    SQLz = SQLz & " AND NOT SUFBack is null"
    Header = "SUF Returned Report " & Me.StartDate & " ~ " & Me.EndDate
    XLS = ""
    XLS = XLS & "|A,DD3|E,BC|F,TC|G,TX|K,SR|H,TC|J,TC|L,SR|M,TC|N,TC|O,BC|P,SR|Q,BC|R,SR|S,SR|T,SR|U,$E"
    XLS = XLS & "|A,W20|E,W09|K,W10|H,W10|J,W10|L,W10|M,W10|N,W10|O,W06|P,W10|Q,W07|S,W10|T,W10|U,W11"
    XLS = XLS & "|E,cColor:3:black|N,cColor:3:green|R,cColor:3:purple|U,cColor:3:crimson"
    XLS = XLS & "|E,cBorder:Med   |N,cBorder:Med   |R,cBorder:Med    |U,cBorder:Med"
    XLS = XLS & "|R,FxAll=P4-K4|R,W10|"
    MessageUser "Open", "Creating SUF Returned Report", "...Getting Data..."
    Call SendReport2Excel(SQLz, "SUF Returned", Header, True, XLS, 0, 1)
    MessageUser "Close"
End Sub
Private Sub MaintInvoiceRequest_Click()
    'DoCmd.Minimize
    DoCmd.OpenForm "BuyersForm", , , , , , Me.Form.Name & ":BC"
End Sub


Private Sub MgtReport_Click()
    Dim XLS, Header As String
    Dim UnitCnt As Integer
    MessageUser "Open", "Creating Future Inventory Report", "...Getting Data..."

    SQLz = "SELECT UnitID, SourceCompany, splitstring(R.[Sch|Grp],'|',1) & '|' & splitstring(R.[Sch|Grp],'|',2) " & _
        "as [Sch|Grp], UnitDescShortVer as UnitDesc, UnitStatus As Status, DateDiff('d',Date(),[OffLeaseDate]) as DaysToEOL, OffLeaseDate as OffLease, Residual, FMVBase AS WhSale, FMV as Retail, TargetSalePrice AS Target, ParkDateProjected AS ProjPark, " & _
        "InspectedDate as InspDate, InspectionPassed as InspPass, InspectionLinkWeb as InspLink " & _
        "From vw_RemarketingReportAllUnits As R " & _
        "WHERE UnitStatus = 'TR' OR UnitStatus = 'P' " & _
        "ORDER BY Offleasedate, SourceCompany, [Sch|Grp];"

    Header = "Future Remarketing Inventory " & Date
    XLS = "|D,TX|E,BC|F,TC|H,$E|I,$R|J,$R|K,$R|L,BL|A,LG1|O,HY|"
    SendReport2Excel SQLz, "Future", Header, True, XLS, UnitCnt, False, 1
    MessageUser "Update", "Creating Current Inventory Report", "...Getting Data..."

    SQLz = "SELECT UnitID, SourceCompany, splitstring(R.[Sch|Grp],'|',1) & '|' & splitstring(R.[Sch|Grp],'|',2) " & _
        "as [Sch|Grp], UnitDescShortVer as UnitDesc, UnitStatus as Status, IIf([DaysInInventory]<0,0,[DaysInInventory]) AS DaysIn, OffLeaseDate as OffLease, Residual, BuyOut, FMVBase AS WhSale, FMV as Retail, TargetSalePrice AS Target, IFBank, " & _
        "InspectedDate as InspDate, InspectionPassed as InspPass, InspectionLinkWeb as InspLink " & _
        "From vw_RemarketingReportAllUnits As R " & _
        "WHERE (Unitstatus = 'R' Or unitstatus = 'PS') " & _
        "ORDER BY UnitStatus DESC , SourceCompany;"

    Header = "Current Remarketing Inventory " & Date
    XLS = "|D,TX|E,BC|F,TC|H,$E|I,$E|J,$R|K,$R|L,$R|M,BL|A,LG1|P,HY|"
    SendReport2Excel SQLz, "Current", Header, True, XLS, UnitCnt, False, 2, True
    MessageUser "Close"
    Set RST = Nothing
End Sub

Private Sub OutstandingPayOffs_Click()
    SQLz = "SELECT BankShortName as Bank, ClientShortName AS Client, Schedule, UnitGroup AS [Group], AssetID, Unitnum, VIN, UnitDescShortVer, [Status], OffLeaseActual"
    SQLz = SQLz & ", BuyOutDate, BuyOutCost, PayOffRequested, PayOffReceived"
    SQLz = SQLz & " FROM vw_SixKeys INNER JOIN Rmkt ON vw_SixKeys.AssetID = Rmkt.UnitRef "
    SQLz = SQLz & " WHERE ((Not (Rmkt.PayOffRequested) Is Null) AND ((Rmkt.PayOffReceived) Is Null)) Order By BankShortName, ClientShortName, AssetID;"
    XLS = "|A,BL|E,TX|E,TL|I,BC|L,$R|"
    Call MessageUser("OPEN", "Getting Report", "Just a moment...")
    Call SendReport2Excel(SQLz, "Units", "OutStanding Pay-Off Request", True, XLS, 0, True)
    Call MessageUser("Close")
End Sub



Private Sub ScheduleControl_Click()
    DoCmd.close acForm, "ScheduleControl"
    DoCmd.OpenForm "ScheduleControl"
End Sub

Private Sub SPDashboard_Click()
    Dim OpenArgX As String
    If FAUserAccess > 9 Then OpenArgX = "Manager"
    If FAUserAccess < 9 Then OpenArgX = "Sales"
    If Me.OverRideCbo = "Sales" Then OpenArgX = "Sales"
    DoCmd.close acForm, "RemarketingSalesPersonsDashBoard"
    DoCmd.OpenForm "RemarketingSalesPersonsDashBoard", , , , , , OpenArgX
    DoCmd.SelectObject acForm, "RemarketingSalesPersonsDashBoard"
End Sub
Private Sub OpenDashBoardInv_Click()
    Call MakeDataSet(2)
    DoCmd.close acForm, "RemarketingDashboardInvSales"
    DoCmd.OpenForm "RemarketingDashboardInvSales"
End Sub
Private Sub OpendashBoardInvMgr_Click()
    If FAUserAccess > 9 Then
        Call MakeDataSet(2)
        DoCmd.close acForm, "RemarketingDashboardInvMgr"
        DoCmd.OpenForm "RemarketingDashboardInvMgr"
    Else
        MsgBox "Sorry, You're not a Manager" & vbNewLine & "Form will not open", vbCritical, "Access Level"
        Exit Sub
    End If
End Sub

Private Sub OpenDashBoardSales_Click()
    If FAUserAccess < 9 Or Me.OverRideCbo = "Sales" Then
        Call MakeDataSet(2)
        DoCmd.close acForm, "RemarketingDashboardSOLDales"
        DoCmd.OpenForm "RemarketingDashboardSOLDSales"
        Exit Sub
    End If
    If FAUserAccess > 9 Or Me.OverRideCbo = "Manager" Then
        Call MakeDataSet(2)
        DoCmd.close acForm, "RemarketingDashboardSOLDMgr"
        DoCmd.OpenForm "RemarketingDashboardSOLDMgr"
        Exit Sub
    End If
End Sub


Private Sub OpenDashBoardAdmin_Click()
    If FAUserAccessLevel > 11 Then
        Call MakeDataSet(2)
        DoCmd.close acForm, "RemarketingDashboardAdmin"
        DoCmd.OpenForm "RemarketingDashboardAdmin"
    Else
        MsgBox "Sorry, You do not have access to that form", , "Security"
    End If
End Sub
Private Sub OpenDashBoardAcct_Click()
    If FAUserAccessLevel > 14 Then
        Call MakeDataSet(2)
        DoCmd.close acForm, "RemarketingDashboardAccounting"
        DoCmd.OpenForm "RemarketingDashboardAccounting"
    Else
        MsgBox "Sorry, You do not have access to that form", , "Who do you think you are?"
    End If
End Sub
Private Sub OpenRMKTMassMod_Click()
    DoCmd.OpenForm "RemarketingMassMod"
End Sub

Private Sub Ctl_UpdateRMKTData_Click()
    Call MessageUser("Open", "Closing Dashboards", "just a few moments...")
    Dim frm As Form
    For Each frm In Forms
        If frm.Name Like "*RemarketingDashboard*" Then DoCmd.close acForm, frm.Name
    Next
    Call MessageUser("Getting Data", "Closing Dashboards", "Almost Done...")
    Call MakeDataSet(0)
    Call MessageUser("Close")
End Sub
Public Sub MakeDataSet(Hoursx As Integer)
    If DateDiff("h", DLookup("LastWriteDate", "LastWriteReport"), Now()) >= Hoursx Then
        DoCmd.SetWarnings False
        DoCmd.RunSQL ("SELECT #" & Now() & "# As LastWriteDate INTO [LastWriteReport];")
        Call SetRemarketingAllUnits
        DoCmd.SetWarnings True
    End If
End Sub

Private Sub SurrenderPDReport_Click()
    Dim SQLz, XLS As String
    Call MessageUser("Open", "Getting Data...", "Excel will open in a bit.")
    SQLz = "SELECT Client, Schedule, Group, UnitID, Unitnum, BuyOutCost, BuyOutDate, EffectiveBuyOut,InspectionPassed, Days, HoldOver, TotalHoldOver, ActualOffLease, UnitStatus as Status "
    SQLz = SQLz & "FROM p_SurrenderPerDiemReport "
    XLS = "|B,BL|C,BL|D,TX|E,TC|F,$E|J,BC|K,$R|L,$E|N,BC|"
    Call SendReport2Excel(SQLz, "HoldOver", "Surrender Per-Diem Report", True, XLS, 0, True)
    Call MessageUser("Close")
End Sub

Private Sub WrongStatus_Click()
    If FAUserAccess < 15 Then
        MsgBox "Sorry you do not have access", vbInformation, "Freeze!"
        Exit Sub
    End If
    DoCmd.OpenForm "RemarketWrongStatus"
End Sub
