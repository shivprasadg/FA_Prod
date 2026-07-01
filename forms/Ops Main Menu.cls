Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
    Dim SQLz As String
    Dim XLS As String
    Dim Headerx As String
    Dim Rowi As Integer
    Dim UnitCnt As Integer

Private Sub AMMenu_Click()
    DoCmd.close acForm, "AM_Menu"
    DoCmd.OpenForm "AM_Menu"
End Sub

Private Sub CloseMe_Click()
    DoCmd.OpenForm "FAMainMenu"
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub RequestPapaerTitles_Click()
    DoCmd.close acForm, "ElectronciTitlePopUp"
    DoCmd.OpenForm "ElectronciTitlePopUp"
End Sub

Private Sub CustDelvReportFunding_Click()
    Call CustDelvReportFinal(1)
End Sub

Private Sub CustodianDashboard_Click()
    DoCmd.OpenForm "TitleToCustodianDashboard"
End Sub

Private Sub CustomerDeliveryReport_Click()
    DoCmd.OpenForm "CustDelSel"
End Sub

Private Sub Customer_ST_Click()
    If IsNull(Me.SelectState) Then
        MsgBox "You must select a State before you can run this report", vbCritical, "What are you doing?!"
        Exit Sub
    End If

    SQLz = "SELECT ClientShNm As Client, LocationName AS Location, LaddrLine1 as Address "
    SQLz = SQLz & ", LocationStreet AS [Co Name], LCity AS City, LCounty AS County"
    SQLz = SQLz & ", LState AS St, LZip AS Zip, C.ClientGroupID "
    SQLz = SQLz & "FROM Clients as C INNER JOIN Locations as L ON C.ClientGroupID = L.ClientGroupID "
    SQLz = SQLz & "WHERE L.LState='" & Me.SelectState & "' "
    SQLz = SQLz & "AND L.LocationType='garage' "
    SQLz = SQLz & "AND (C.clienttype='client' Or C.clienttype='guarantor') "
    SQLz = SQLz & "AND L.locobsolete=No;"
    Call SendReport2Excel(SQLz, "Clients by State", "Clients By State", True, , , True)
End Sub

Private Sub LocationAddrManager_Click()
    If Me.SelectClient & "test" = "test" Then
            MsgBox "You must select a Client Group before you can run this Form", vbCritical, "What are you doing?!"
        Exit Sub
    End If
        DoCmd.OpenForm "LocationAddressesEdit", acNormal, , "LocationID=" & DMin("LocationID", "Locations", "LocationType='Garage' AND ClientGroupID=" & Me.SelectClient) & ""
End Sub

Private Sub MissingPDStartDates_Click()
    SQLz = "Select"
    SQLz = SQLz & " ClientShortName As Client"
    SQLz = SQLz & ", Schedule"
    SQLz = SQLz & ", UnitGroup"
    SQLz = SQLz & ", FundedDate"
    SQLz = SQLz & ", AssetID"
    SQLz = SQLz & ", VIN"
    SQLz = SQLz & ", UnitVendorPytDate as VendorPaid"
    SQLz = SQLz & ", UnitPDStartDate as [PD Start]"
    SQLz = SQLz & ", UNITPDRENT as [PD Rent]"
    SQLz = SQLz & " From vw_SixKeys As S"
    SQLz = SQLz & " Inner Join Units as U on S.AssetID=U.UnitID"
    SQLz = SQLz & " Where Not FundedDate Is Null"
    SQLz = SQLz & " AND FundedDate > #1/1/2023#"
    SQLz = SQLz & " AND U.UnitPDStartDate Is Null"
    SQLz = SQLz & " ORDER BY ClientShortName, Schedule, UnitGroup, VIN"
    XLS = "|F,TC|D,TC|E,BC|G,TC|I,$E|"
    Call SendReport2Excel(SQLz, "Needs PD Start", "Units that need PD Start", True, XLS, , True)
End Sub
Private Sub FundedNotDelivered_Click()
    Call MessageUser("Open", "Getting Delivered - Vendor Not Paid Report", "...This will take moment...")
    SQLz = "SELECT ClientShortName As Client"         'A
    SQLz = SQLz & ", BDEXDeal as BDEX"                'B
    SQLz = SQLz & ", Schedule"                        'C
    SQLz = SQLz & ", UnitGroup"                       'D
    SQLz = SQLz & ", AssetID"                         'E
    SQLz = SQLz & ", UnitDescAbbr as UnitDesc"        'F
    SQLz = SQLz & ", VendorSNm as Vendor"             'G
    SQLz = SQLz & ", UnitInvoiceDate as [Inv Date]"   'H
    SQLz = SQLz & ", DateDiff('d',UnitInvoiceDate,nz(VendorPaidDate,Date())) AS [Days Inv]" 'I
    SQLz = SQLz & ", VendorPaidDate as [Vendor Paid]" 'J
    SQLz = SQLz & ", DateDiff('d',VendorPaidDate,Date()) AS [Days Paid]" 'K
    SQLz = SQLz & ", ActualDeliveryDate as [Delv Date]" 'L
    SQLz = SQLz & ", DateDiff('d',ActualDeliveryDate,Date()) AS [Days Delv]" 'M
    SQLz = SQLz & ", [Interim Bank]"                  'N
    SQLz = SQLz & ", [Interim Amt]"                   'O
    SQLz = SQLz & ", BankShortName as Assignee"       'P
    SQLz = SQLz & ", FundDate"                        'Q
    SQLz = SQLz & " FROM vw_CustDelvReportBase"
    SQLz = SQLz & " WHERE NOT FundDate Is Null "
    SQLz = SQLz & " AND ActualDeliveryDate Is Null "
    SQLz = SQLz & " ORDER BY ClientShortName, Schedule;"
    Call MessageUser("Update", "Updating Delivered - Not Funded Report", "...Excel will open soon...")
    Headerx = "Funded NOT Delivered: " & Date
    XLS = "|A,DD|A,Wrap|B,TC|B,W7|E,TX|E,BC|H,TC|H,TC|I,TC|J,W10|K,Sint|L,BC|M,TC|M,Sint|M,BC|O,$E|"
    XLS = XLS & "B,cColor:3:0.0.0|"
    XLS = XLS & "N,cColor:3:0.72.86|O,cColor:3:0.72.86|"
    XLS = XLS & "J,cColor:3:153.153.0|K,cColor:3:153.153.0|"
    XLS = XLS & "L,cColor:3:0.0.0|M,cColor:3:0.0.0|"
    XLS = XLS & "G,cLabel:End=J:2:" & TLeftLine & LineSpace(10) & " Vendor Details " & LineSpace(25) & TRightLine & "|"
    XLS = XLS & "L,cLabel:End=M:2:" & TLeftLine & LineSpace(2) & " Delivery " & LineSpace(10) & TRightLine & "|"
    SendReport2Excel SQLz, "Not Funded", Headerx, True, XLS, 0, True
    Call MessageUser("Close")
End Sub

Private Sub LocationManager_Click()
 DoCmd.close acForm, "LocationTaxManager"
 DoCmd.OpenForm "LocationTaxManager"
End Sub

Private Sub OpenAuditManager_Click()
DoCmd.close acForm, "AuditManager"
DoCmd.OpenForm "AuditManager"
End Sub

Private Sub OPSMetrics_Click()
    Call MessageUser("Open", "Getting Metrics Report", "...This will be quick...")
    Call getOPSMetricsReport(Nz(Me.StartDate, ""), Nz(Me.EndDate, ""))
    Call MessageUser("Close")
End Sub

Private Sub PayProceedsManager_Click()
    DoCmd.close acForm, "PayProceedsManager"
    DoCmd.OpenForm "PayProceedsManager"
End Sub

Private Sub RegLocations_Click()
If Me.SelectClient & "test" = "test" Then
MsgBox "You must select a Client Group" & vbNewLine & "before you can run this report", vbCritical, "What are you doing?!"
Exit Sub
End If
   Call MessageUser("Open", "Getting REG Locations", "...This will take moment...")
    SQLz = "Select Client"
    SQLz = SQLz & ",Schedule"
    SQLz = SQLz & ",UnitGroup"
    SQLz = SQLz & ",AssetID"
    SQLz = SQLz & ",Status"
    SQLz = SQLz & ",UnitNum"
    SQLz = SQLz & ",VIN"
    SQLz = SQLz & ",UnitDesc"
    SQLz = SQLz & ",RegLocation"
    SQLz = SQLz & ",RegAddress"
    SQLz = SQLz & ",RState"
    SQLz = SQLz & ",RCounty"
    SQLz = SQLz & ",TaxOption"
    SQLz = SQLz & ",TaxRate"
    SQLz = SQLz & ",LocationID"
    SQLz = SQLz & ",ClientID"
    SQLz = SQLz & ",ClientGroupID"
    SQLz = SQLz & " FROM vw_RegLocations "
    SQLz = SQLz & " WHERE ClientGroupId=" & Me.SelectClient & " "
    SQLz = SQLz & " Order By Client, RegLocation"
    XLS = "|A,Wrap|A,zCenter|A,TL|C,TX|C,BC|E,TC|J,W50|K,BC|M,P3|A,Banding|O,HH|P,HH|Q,HH|"
    Call SendReport2Excel(SQLz, "Reg Locations", "Reg Locations by Client", True, XLS)
    Call MessageUser("Close")
End Sub

Private Sub TitleByState_Click()
    If IsNull(Me.SelectState) Then
        MsgBox "You must select a State before you can run this report", vbCritical, "What are you doing?!"
        Exit Sub
    End If
    SQLz = "Select Client"
    SQLz = SQLz & ",MLNo"
    SQLz = SQLz & ",Schedule"
    SQLz = SQLz & ",[Group]"
    SQLz = SQLz & ",AssetID"
    SQLz = SQLz & ",VIN"
    SQLz = SQLz & ",TitleNumber"
    SQLz = SQLz & ",TitleSt"
    SQLz = SQLz & ",EntryReason"
    SQLz = SQLz & ",UnitTitleLogDate As TitleLogDate"
    SQLz = SQLz & ",UnitTitleLogId "
    SQLz = SQLz & "From UnitTitleHistoryLastEntry "
    SQLz = SQLz & "WHERE TitleSt ='" & Me.SelectState & "' "
    SQLz = SQLz & "Order By Client "
    XLS = "|A,BL|E,TX|E,B|H,BC|K,HH|"
    Call SendReport2Excel(SQLz, "Titles In " & Me.SelectState, "Titles By State: " & Me.SelectState, True, XLS, , True)
End Sub
Private Sub CustomerDeliveryReportDetailed_Click()
    DoCmd.OpenReport "CustomerDeliveryReportDetailed", acViewPreview
End Sub
Private Sub DeliveredNotFunded_Click()
    Call MessageUser("Open", "Getting Delivered - Vendor Not Paid Report", "...This will take moment...")
    SQLz = "SELECT ClientShortName As Client"         'A
    SQLz = SQLz & ", BDEXDeal as BDEX"                'B
    SQLz = SQLz & ", Schedule"                        'C
    SQLz = SQLz & ", UnitGroup"                       'D
    SQLz = SQLz & ", AssetID"                         'E
    SQLz = SQLz & ", UnitDescAbbr as UnitDesc"        'F
    SQLz = SQLz & ", VendorSNm as Vendor"             'G
    SQLz = SQLz & ", UnitInvoiceDate as [Inv Date]"   'H
    SQLz = SQLz & ", DateDiff('d',UnitInvoiceDate,nz(VendorPaidDate,Date())) AS [Days Inv]" 'I
    SQLz = SQLz & ", VendorPaidDate as [Vendor Paid]" 'J
    SQLz = SQLz & ", DateDiff('d',VendorPaidDate,Date()) AS [Days Paid]" 'K
    SQLz = SQLz & ", ActualDeliveryDate as [Delv Date]" 'L
    SQLz = SQLz & ", DateDiff('d',ActualDeliveryDate,Date()) AS [Days Delv]" 'M
    SQLz = SQLz & ", [Interim Bank]"                  'N
    SQLz = SQLz & ", [Interim Amt]"                   'O
    SQLz = SQLz & ", BankShortName as Assignee"       'P
    SQLz = SQLz & ", FundDate"                        'Q
    SQLz = SQLz & " FROM vw_CustDelvReportBase"
    SQLz = SQLz & " WHERE FundDate Is Null "
    SQLz = SQLz & " AND NOT ActualDeliveryDate Is Null "
    SQLz = SQLz & " ORDER BY ClientShortName, Schedule;"
    Call MessageUser("Update", "Updating Delivered - Not Funded Report", "...Excel will open soon...")
    Headerx = "Delivered, Not Funded: " & Date
    XLS = "|A,DD|A,Wrap|B,TC|B,W7|E,TX|E,BC|H,TC|H,TC|I,TC|J,W10|K,Sint|L,BC|M,TC|M,Sint|M,BC|O,$E|"
    XLS = XLS & "B,cColor:3:0.0.0|"
    XLS = XLS & "N,cColor:3:0.72.86|O,cColor:3:0.72.86|"
    XLS = XLS & "J,cColor:3:153.153.0|K,cColor:3:153.153.0|"
    XLS = XLS & "L,cColor:3:0.0.0|M,cColor:3:0.0.0|"
    XLS = XLS & "G,cLabel:End=J:2:" & TLeftLine & LineSpace(10) & " Vendor Details " & LineSpace(25) & TRightLine & "|"
    XLS = XLS & "L,cLabel:End=M:2:" & TLeftLine & LineSpace(2) & " Delivery " & LineSpace(10) & TRightLine & "|"
    SendReport2Excel SQLz, "Not Funded", Headerx, True, XLS, 0, True
    Call MessageUser("Close")
End Sub
Private Sub CustomerDeliveryReportNew_Click()
 Call CustDelvReportFinal
End Sub


Private Sub CustomerDeliveryReportSummary_Click()
    SQLz = "Select * From vw_CustomerDeliveryReportSummary;"
    Headerx = "Customer Delivery Report Summary - " & Format(Date, "mm-dd-yy")
    XLS = "|B,TX|D,TC|J,TC|L,TC|N,TC|Q,$E|P,TC|S,TR|T,ZZ|T,ZZ|"
    SendReport2Excel SQLz, "CDReport", Headerx, True, XLS, DCount("SchID", "vw_CustomerDeliveryReportSummary"), True
End Sub

Private Sub DeliveredUnitsNotInService_Click()

    Dim RST As DAO.Recordset
    SQLz = ""
    SQLz = SQLz & " SELECT ClientShortName"
    SQLz = SQLz & ", BDEX "
    SQLz = SQLz & ", Schedule.SchShortDesc "
    SQLz = SQLz & ", Schedule"
    SQLz = SQLz & ", UnitGroup "
    SQLz = SQLz & ", AssetID, Status "
    SQLz = SQLz & ", Unitnum, VIN, UnitDescShortVer AS UnitDesc"
    SQLz = SQLz & ", ActualDeliveryDate "
    SQLz = SQLz & ", DateDiff('d',[ActualDeliveryDate],Date()) AS DeliveredDays "
    SQLz = SQLz & ", AcceptedDate "
    SQLz = SQLz & ", DateDiff('d',[AcceptedDate],Date()) AS AcceptedDays "
    SQLz = SQLz & ", InServiceDate"
    'Sqlz = Sqlz & ",Mlorig, clientID, ClientGroupId, vw_SixKeys.SchID, vw_SixKeys.SGrpID "
    SQLz = SQLz & " FROM vw_SixKeys INNER JOIN Schedule ON vw_SixKeys.SchID = Schedule.SchID "
    SQLz = SQLz & " WHERE (((Mlorig)='FA') "
    SQLz = SQLz & " AND ((Status)<>'C') "
    SQLz = SQLz & " AND ((ClientGroupId)<>445) "
    SQLz = SQLz & " AND ((ActualDeliveryDate) Is Not Null) "
    SQLz = SQLz & " AND ((InServiceDate) Is Null "
    SQLz = SQLz & " Or (InServiceDate)>Date())) "
    SQLz = SQLz & " ORDER BY ClientShortName, vw_SixKeys.SchID, UnitGroup, AssetID;"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    RST.MoveLast
    Headerx = "Units To Be In-Serviced Excel - " & Format(Date, "mm-dd-yy")
    XLS = "|A,TL|B,BC|E,TC|F,TX|G,BC|L,BC|N,BC|"
    SendReport2Excel SQLz, "CDReport", Headerx, True, XLS, RST.RecordCount, True
    RST.close
    Set RST = Nothing
End Sub

Private Sub DelvNotInService_Click()

   SQLz = "SELECT S.CompanyName"
   SQLz = SQLz & " ,S.MLNo"
   SQLz = SQLz & " ,S.FAID"
   SQLz = SQLz & " ,S.Schedule"
   SQLz = SQLz & " ,S.UnitGroup AS [Group]"
   SQLz = SQLz & " ,S.AssetID"
   SQLz = SQLz & " ,S.Unitnum"
   SQLz = SQLz & " ,S.VIN"
   SQLz = SQLz & " ,S.ActualDeliveryDate AS [Actual Delv Date]"
   SQLz = SQLz & " ,S.AcceptedDate AS [Accepted Date]"
   SQLz = SQLz & " ,S.InServiceDate AS [In Service Date]"
   SQLz = SQLz & " ,U.UnitCOASent as [CoOA Sent]"
   SQLz = SQLz & " ,U.UnitCOABack AS [CoA Back]"
   SQLz = SQLz & " ,S.Status"
   SQLz = SQLz & " ,S.UnitDocAccpt AS [Doc Accept]"
   SQLz = SQLz & " ,S.UnitDescShortVer AS [Desc]"
   SQLz = SQLz & " ,Date() - [ActualDeliveryDate] AS [Days In Use]"
   SQLz = SQLz & " ,U.UnitPDStartDate AS [PD Start Date] "
   SQLz = SQLz & " ,U.UNITPDRENT AS [PD Rent]"
   SQLz = SQLz & " FROM vw_SixKeys AS S"
   SQLz = SQLz & " INNER JOIN Units As U ON S.AssetID = U.UnitID"
   SQLz = SQLz & " WHERE S.ActualDeliveryDate IS NOT NULL "
   SQLz = SQLz & " AND   S.InServiceDate IS NULL "
   SQLz = SQLz & " AND   S.Status in ('A','B')"
   SQLz = SQLz & " ORDER BY S.CompanyName;"
   XLS = "|A,mTable:Name=Delivered|A,DD3|A,BL|D,TC|E,TC|F,TX|F,BC|F,cBorder:Med|F,cColor:3:Black|G,TC|I,TC|J,TC|K,TC|L,TC"
   XLS = XLS & "|M,TC|N,BC|N,cBorder:Med|N,cColor:3:Black|Q,BC|R,TC|R,cColor:3:Green|S,$R|S,cBorder:Med|S,cColor:3:Green|K,cColor:3:Black|K,cFontColor:3:Yellow"
   XLS = XLS & "|A,W25|I,W10|J,W10|K,W10|L,W10|M,W10|O,W08|Q,W06|R,W10|"
   MessageUser "Open", "Creating Delv, Not Inservice Report, That's Tab 1", "...Excel Will Open in a bit..."
   TitleX = "Units Delivered Not Inservice | " & Date
   SendReport2Excel SQLz, "Not In Service", TitleX, True, XLS, , False, 1 ', , , "C:\Faas\" & FAUserName & "\Temp\UnitsDeliveredNotInService.xlsx"
   SQLz = "SELECT S.CompanyName"
   SQLz = SQLz & " ,S.MLNo"
   SQLz = SQLz & " ,S.FAID"
   SQLz = SQLz & " ,S.Schedule"
   SQLz = SQLz & " ,S.UnitGroup AS [Group]"
   SQLz = SQLz & " ,S.AssetID"
   SQLz = SQLz & " ,S.Unitnum"
   SQLz = SQLz & " ,S.VIN"
   SQLz = SQLz & " ,S.ActualDeliveryDate AS [Actual Delv Date]"
   SQLz = SQLz & " ,S.AcceptedDate AS [Accepted Date]"
   SQLz = SQLz & " ,S.InServiceDate AS [In Service Date]"
   SQLz = SQLz & " ,U.UnitCOASent as [CoOA Sent]"
   SQLz = SQLz & " ,U.UnitCOABack AS [CoA Back]"
   SQLz = SQLz & " ,S.Status"
   SQLz = SQLz & " ,S.UnitDocAccpt AS [Doc Accept]"
   SQLz = SQLz & " ,S.UnitDescShortVer AS [Desc]"
   SQLz = SQLz & " ,Date() - [ActualDeliveryDate] AS [Days In Use]"
   SQLz = SQLz & " ,U.UnitPDStartDate AS [PD Start Date] "
   SQLz = SQLz & " ,U.UNITPDRENT AS [PD Rent]"
   SQLz = SQLz & " FROM vw_SixKeys AS S"
   SQLz = SQLz & " INNER JOIN Units As U ON S.AssetID = U.UnitID"
   SQLz = SQLz & " WHERE S.AcceptedDate IS NOT NULL "
   SQLz = SQLz & " AND   U.UnitPDStartDate IS NULL "
   SQLz = SQLz & " AND   S.Status in ('A','B')"
   SQLz = SQLz & " ORDER BY S.CompanyName;"
   MessageUser "Open", "Creating Units Accepted No PerDiem Report, That's Tab 2", "...Excel Will Open in a bit..."
   TitleX = "Units Accepted No PerDiem | " & Date
   SendReport2Excel SQLz, "No PerDiem", TitleX, True, XLS, , False, 2, True
   MessageUser "Close"
 End Sub

Private Sub DrawManager_Click()
   DoCmd.close acForm, "DrawDownManager"
   DoCmd.OpenForm "DrawDownManager"
End Sub

Private Sub ExitOptionsConsole_Click()
    Call MessageUser("Open", "Opening Exit Options form", "...This will take just a moment...")
    DoCmd.close acForm, "ExitOptionsManager"
    DoCmd.OpenForm "ExitOptionsManager"
    Call MessageUser("Close")
    Forms!ExitOptionsManager.SetFocus
End Sub

Private Sub FundedUnitReport_Click()
    If IsNull(Me.StartDate) Or IsNull(Me.EndDate) Then
        MsgBox "You must select a BOTH a Start and End date" & vbNewLine & "before you can run this report", vbCritical, "What are you doing?!"
        Exit Sub
    End If

    SQLz = "SELECT S.ClientShortName AS Client, S.Schedule AS Sch, S.UnitGroup AS Grp, S.AssetID, S.UnitNum, S.VIN, S.FundedDate, S.UnitInvoiceNum AS InvoiceNum"
    SQLz = SQLz & ", S.UnitInvoiceDate AS InvoiceDate, S.VendorCost, UP.VendorPaid, TA.TitleAppSent, DrawName, DrawNumber AS DrawNum, UP.UnitIfBank, UP.UnitIFAmt AS IFAmount "
    SQLz = SQLz & " FROM ((Vw_SixKeys AS S "
    SQLz = SQLz & " INNER JOIN (Select UnitID, UnitVendorPytDate as VendorPaid, DrawDownId, UnitIFAmt, BankShortName as UnitIfBank "
    SQLz = SQLz & "             From Units"
    SQLz = SQLz & "             Left Join Banks ON Units.UnitIFBankID=Banks.BankID) AS UP ON S.AssetID = UP.UnitID)"
    SQLz = SQLz & " LEFT JOIN (Select UnitRef, UnitTitleLogDate as TitleAppSent, EntryReason From UnitTitleLog Where EntryReasonId = 18) AS TA ON S.AssetID=TA.UnitRef)"
    SQLz = SQLz & " LEFT JOIN DrawDowns ON UP.DrawDownId = DrawDowns.DrawDownId"
    SQLz = SQLz & " Where VendorPaid >= #" & Me.StartDate & "# AND VendorPaid <= #" & Me.EndDate & "# "
    SQLz = SQLz & " Order By ClientShortName, Schedule, UnitGroup, VIN"
    XLS = "|D,BC|G,BC|G,BC|J,$E|N,BC|P,$E|"
    SendReport2Excel SQLz, "Funded Units", "Funded Units: " & Me.StartDate & "~" & Me.EndDate, True, XLS, , True

End Sub

Private Sub MakeModelManager_cmd_Click()
    DoCmd.close acForm, "MakeModelManager"
    DoCmd.OpenForm "MakeModelManager"
End Sub
Private Sub ChildPartManagement_Click()
Exit Sub
    DoCmd.close acForm, "MakeModelManagerParts"
    DoCmd.OpenForm "MakeModelManagerParts"
End Sub
Private Sub MissingPlateNumber_Click()

    XLS = "|A,BL|D,TC|G,BC|H,BC|M,TR|N,BC|"
    UnitCnt = DCount("AssetId", "UnitsMissingLicensePlates")
    FormatExcel "UnitsMissingLicensePlates", "C:\Faas\" & FaUserName & "\Temp\UnitsMissingLicensePlates.xlsx", "MissingPlateNum", , XLS, Countx, "Missing Plate Numbers | " & UnitCnt & " units", True

End Sub


Private Sub Form_Load()
'If Forms!FAMainMenu.USERID = "ssweet" Or Forms!FAMainMenu.USERID = "ljohnson" Then Me.AM_Menu.Visible = True Else Me.AM_Menu.Visible = False
End Sub

Private Sub MassModForm_Click()
    DoCmd.close acForm, "BulkChange_Form"
    DoCmd.OpenForm "BulkChange_Form"
End Sub

Private Sub OpenBankView_Click()
    DoCmd.close acForm, "Bank_View"
    DoCmd.OpenForm "Bank_View"
End Sub

Private Sub OpenReLeaseManage_Click()
    DoCmd.close acForm, "ReLeaseUnitManager"
    DoCmd.OpenForm "ReLeaseUnitManager"
End Sub

Private Sub OpenUnitLookUp_Click()
    DoCmd.OpenForm "UnitReverseLookUpForm"
End Sub

Private Sub OpenVendorForm_Click()
    DoCmd.OpenForm "VendorsForm"
End Sub

Private Sub OutstandingExcel_Click()
SQLz = "Select * From OutstandingTitles"
XLS = "|B,TC|B,W10|F,CX|F,TC|G,TC|I,BC|J,TC|M,TC|N,Sint|N,BC|J,cColor:3:153.153.0|N,cColor:3:153.153.0|O,CX|"
SendReport2Excel SQLz, "Titles Out", "Outstanding Titles: " & Me.StartDate & "~" & Me.EndDate, True, XLS, , True
'FormatExcel "OutstandingTitles", "C:\Faas\" & FAUserName & "\Temp\OutstandingTitles", "Titles Out", , "|B,TC|F,CX|G,TC|I,BC|M,TC|N,Sint|N,BC|", DCount("AssetID", "OutstandingTitles", ""), "Outstanding Titles, as of: " & Date, True
End Sub



Private Sub OutstandingSUFsExcel_Click()
    Call Remarketing_OutStanding_SUF
End Sub

Private Sub Print_PDF_SUF_Click()
    DoCmd.OpenForm "SUF_Select_Form"
End Sub



Private Sub RMKT_Title_Sent_Click()
    If IsNull(Me.StartDate) Or IsNull(Me.EndDate) Then
        MsgBox "You must select a BOTH a Start and End date" & vbNewLine & "before you can run this report", vbCritical, "What are you doing?!"
        Exit Sub
    End If

    SQLz = "SELECT ClientShortName as Client, Count(UnitRef) AS Units, Min(UnitTitleLogDate) as First, Max(UnitTitleLogDate) as Last "
    SQLz = SQLz & "FROM vw_SixKeys As Six "
    SQLz = SQLz & "INNER JOIN UnitTitleLog AS TL ON Six.UnitID = TL.UnitRef "
    SQLz = SQLz & "WHERE TL.EntryReasonID = 7 "
    SQLz = SQLz & "AND UnitTitleLogDate >= #" & Me.StartDate & "# "
    SQLz = SQLz & "AND UnitTitleLogDate <= #" & Me.EndDate & "# "
    SQLz = SQLz & "GROUP BY ClientShortName "
    SQLz = SQLz & "ORDER BY ClientShortName;"

    SendReport2Excel SQLz, "RMKT TitleSent", "RMKT Titles Sent: " & Me.StartDate & "~" & Me.EndDate, True, , , True
   End Sub

Private Sub TitleAppStatus_Click()
    If IsNull(Me.StartDate) Then
        MsgBox "You must select a Start date" & vbNewLine & "before you can run this report", vbCritical, "What are you doing?!"
        Exit Sub
    End If
    SQLz = "SELECT * from vw_TitleApp_Status "
    Headerx = "Title Apps Status On or after: " & Me.StartDate
    If Nz(Me.SelectClient, 0) > 0 Then
        SQLz = SQLz & "WHERE ClientGroupID=" & Me.SelectClient & " AND (OutDate >= #" & Me.StartDate & "# OR OutDate IS Null)"
        Headerx = Headerx & " Client Group: " & Me.SelectClient.Column(1)
    Else
        SQLz = SQLz & "WHERE (OutDate >= #" & Me.StartDate & "# OR OutDate IS Null)"
    End If
    SQLz = SQLz & "ORDER BY Client, AssetID;"
    XLS = "|C,BL|E,LX|F,BC|J,BC|L,BC|O,BC|R,HH|S,HH|"
    SendReport2Excel SQLz, "TitleApps", Headerx, True, XLS, , True
End Sub
Private Sub TitleSent_Click()
    If IsNull(Me.StartDate) Or IsNull(Me.EndDate) Then
         MsgBox "You must select a BOTH a Start and End date" & vbNewLine & "before you can run this report", vbCritical, "What are you doing?!"
        Exit Sub
    End If

    SQLz = "SELECT ClientShortName as Client, Count(U.UnitID) AS Units, Min(UnitTASent) as First, Max(UnitTASent) as Last "
    SQLz = SQLz & "FROM vw_SixKeys As Six "
    SQLz = SQLz & "INNER JOIN Units As U ON Six.UnitID = U.UnitID "
    SQLz = SQLz & "WHERE UnitTASent >= #" & Me.StartDate & "# "
    SQLz = SQLz & "AND UnitTASent <= #" & Me.EndDate & "# "
    SQLz = SQLz & "GROUP BY ClientShortName "
    SQLz = SQLz & "ORDER BY ClientShortName;"
    SendReport2Excel SQLz, "TitleAppSent", "Title Apps Sent: " & Me.StartDate & "~" & Me.EndDate, True, , , True
End Sub

Private Sub RequestMSO_Click()
    DoCmd.OpenForm "ReqInvMSO_Form"
End Sub


Private Sub SyndicationReadiness_dnu_Click()
    DoCmd.OpenReport "SyndReadyReport", acViewPreview
End Sub

Private Sub BillingRules_Click()
SQLz = "Select * "
SQLz = SQLz & " From vw_GroupBillingRules"
SQLz = SQLz & " ORDER BY Client, Sch, Grp"
XLS = "|E,BC|F,TC|G,TC|H,TC|I,TC|J,TC|K,TC|"
Call MessageUser("Open", "Getting Billing Rules Report", "...Excel will open soon...")

SendReport2Excel SQLz, "Rules", "Group Billing Rules as of: " & Date & "", True, XLS, 0
Call MessageUser("Close")
End Sub

Private Sub Trackingsheet_Click()
    DoCmd.OpenForm "Open_Items"
End Sub

Private Sub UnitCountLocation_Click()

    SQLz = "Select * From ParkedLocationsUnitCountByType;"
    Headerx = "In Use Unit Count At Location Report - All Clients - " & Format(Date, "mm-dd-yy")
    XLS = "|E,SE|F,SE|G,SE|H,SE|I,SE|J,SE|K,SE|L,SE|M,SE|Q,SE|R,SE|S,SE|T,SE|U,SE|V,SE|W,SE|X,SE|Z,ZZ|Z,ZZ|"
    SendReport2Excel SQLz, "Location Counts", Headerx, True, XLS, DCount("ClientID", "ParkedLocationsUnitCountByType"), True
End Sub

Private Sub VinDecoder_Click()
    Call OpenGetVINDialog
End Sub
