Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz As String
Private Sub AddExpBtn_Click()
    DoCmd.OpenForm "SgrpExpForm"
End Sub

Private Sub Assignee2Client_Click()
    SQLz = "SELECT S.BankShortName, S.ClientShortName, S.MLNo, S.Schedule, Count(S.UnitID) AS Units " & "From vw_SixKeysAS S GROUP BY S.BankShortName, S.ClientShortName, S.MLNo, S.Schedule " & "HAVING (((S.BankShortName) <> 'TBD')) ORDER BY S.BankShortName, S.ClientShortName, S.MLNo, S.Schedule;"
    Call SendReport2Excel(SQLz, "Assignee-Client", "Assignee to Client", True, XLS, 0, True)
End Sub

Private Sub BatchBillingForm_Click()
    DoCmd.close acForm, "BatchBillingForm"
    DoCmd.OpenForm "BatchBillingForm"
End Sub
Private Sub InterimStatusReport_Click()
      Call RMKTInterimStatusReport
End Sub

Private Sub CasualtyReport_Click()
    MessageUser "Open", "Getting Casualty Report ", "This will be a moment..."
    SQLz = "SELECT ClientShortName as Client, MlNo, Schedule as Sch, UnitGroup as Grp, BankShortName As Assignee,AssetID, UnitNum, VIN,MYear"
    SQLz = SQLz & ",Make,Model,SubType,[Type],[Status],BLCD"
    SQLz = SQLz & ",LeaseTermPrimary as Term,LeaseExp"
    SQLz = SQLz & ",OffLeaseActual as [OffLease Actual]"
    SQLz = SQLz & ",DateDiff(""M"",BLCD,OffLeaseActual) as AgeBLCD"
    SQLz = SQLz & ",CurrRentType as [Curr Rent Type],CurrRent,OECUnit"
    SQLz = SQLz & " From vw_SixKeys as S"
    SQLz = SQLz & " WHERE S.[Status]='C'"
    SQLz = SQLz & " AND NOT OffLeaseActual IS NULL"
    SQLz = SQLz & " ORDER BY ClientShortName, Schedule,OffLeaseActual"
    XLS = "|A,DD|C,W15|D,TC|F,TC|F,TX|I,TC|N,BC|P,BC|R,W11|S,BC|T,TC|T,W10|U,$|V,$intE|"
    AddReportLog "CasualtyReport_Click"
    Call SendReport2Excel(SQLz, "Casualty Report", "Casualty Report: " & Date, True, XLS, 0, False)
    MessageUser "Close"
End Sub

Private Sub CloseMe_Click()
    DoCmd.OpenForm "FAMainMenu"
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub Form_Close()
    DoCmd.OpenForm "FAMainMenu"
End Sub

Private Sub GoHome_Click()
    DoCmd.OpenForm "FAMainMenu"
End Sub

Private Sub InvoiceClients_Click()
    DoCmd.OpenForm "Billing"
End Sub

Private Sub LoanReport_Click()
    DoCmd.OpenForm "LoanReportParameters"
End Sub

Private Sub ManageInvoice_Click()
    'RmktBillSel
    DoCmd.close acForm, "ReMarketInvoiceManager"
    DoCmd.OpenForm "ReMarketInvoiceManager"
End Sub

Private Sub OpenClientsForm_Click()
    DoCmd.OpenForm "ClientsView"
End Sub

Private Sub PotentialBillingExcel_Click()
    MessageUser "Open", "Getting Potential Billing ", "This will be a moment..."
    SQLz = "Select Client"        'A
    SQLz = SQLz & ",Assignee"     'B
    SQLz = SQLz & ",Portfolio"    'C
    SQLz = SQLz & ",Schedule"     'D
    SQLz = SQLz & ",UnitGroup"    'E
    SQLz = SQLz & ",BLCD"         'F
    SQLz = SQLz & ",Assigned"     'G
    SQLz = SQLz & ",FundDateProj" 'H
    SQLz = SQLz & ",FundStatus"   'I
    SQLz = SQLz & ",UStatus"      'J
    SQLz = SQLz & ",Units"        'K
    SQLz = SQLz & ",FAInvoicing"  'L
    SQLz = SQLz & ",FAServicer"   'M
    SQLz = SQLz & ",FACollecting" 'N
    SQLz = SQLz & ",Strips"       'O
    SQLz = SQLz & " From vw_PotentialBilling;"
    XLS = "|A,DD|E,TX|I,cLabel:End=I:3:Fund" & vbNewLine & "Status|H,cLabel:End=H:3:Fund" & vbNewLine & "Date Proj"
    XLS = XLS & "|D,TL|E,cLabel:End=E:3:Unit" & vbNewLine & "Group|E,TC"
    XLS = XLS & "|B,BL|E,TC|I,BC|J,TC|L,TC|M,TC|N,TC|O,BC|"
    AddReportLog "PotentialBillingExcel_Click"
    Call SendReport2Excel(SQLz, "Potential_Billing", "Potential Billing: " & Date, True, XLS, 0, False)
    MessageUser "Close"
End Sub
Private Sub Ctl_InvoicedReport_Click()
    If Me.StartDate & "test" = "test" Or Me.EndDate & "test" = "test" Then
        MsgBox "You must have an BOTH a Start and End Date", vbInformation, "Date Req"
        Exit Sub
    End If
    SQLz = "Select ClientShortName As Client "
    SQLz = SQLz & ",Schedule,BillType,S.BankShortName as Bank,BillNo,Billdt as BillDate"
    SQLz = SQLz & ",Billfrom,Billto,BillDueDt, BillSent,BillID"
    SQLz = SQLz & " From Bills As B"
    SQLz = SQLz & " Inner Join vw_SixKeys as S On B.BillSchRef = S.SchID"
    SQLz = SQLz & " Where Not Billto Is Null"
    SQLz = SQLz & " AND Billfrom >=#" & Me.StartDate & "# AND Billfrom <=#" & Me.EndDate & "# "
    SQLz = SQLz & " Group By  ClientShortName,Schedule,BillType,BankShortName,BillNo,Billdt"
    SQLz = SQLz & ",Billfrom,Billto,BillDueDt,BillSent,BillID"
    SQLz = SQLz & " Order By ClientShortName, Schedule, BillType, Billfrom"
    MessageUser "Open", "Getting Invoiced Report", "This will be a moment..."
    XLS = "|A,BL|"
    AddReportLog "InvoicedReport"
    Call SendReport2Excel(SQLz, "Invoiced", "Invoiced From " & Me.StartDate & " to: " & Me.EndDate, True, XLS, 0, False)
    MessageUser "Close"
End Sub

Private Sub PotentialBillingPDExcel_Click()
    SQLz = "SELECT Client, Schedule, Group, Min(PDStart) AS StartPDBilling"
    SQLz = SQLz & ", Max(BLCD) AS EndPDBilling, "
    SQLz = SQLz & "Round(Avg([TotalPDDays]),0) AS BillibleDays"
    SQLz = SQLz & ", Count(AssetID) AS Units, Avg(RentPD) AS aRentPD"
    SQLz = SQLz & " From vw_PotentialBillingPerDiem"
    SQLz = SQLz & " Where (((Date()) <= DateAdd('d', 90, [PDStart]))) "
    SQLz = SQLz & " GROUP BY Client, Schedule, Group, BDay, SchID, GroupID"
    SQLz = SQLz & " ORDER BY Client, Min(PDStart) DESC , Schedule, Group;"
    MessageUser "Open", "Getting Invoiced Report", "This will be a moment..."
    XLS = "|A,BL|F,BC|G,TC|H,$R|"
    AddReportLog "PotentialBillingPDExcel_Click"
    Call SendReport2Excel(SQLz, "PD Billing", "Future/Recent PD Billing", True, XLS, 0, False)
    MessageUser "Close"
End Sub

Private Sub RemarkingBuyOut_Click()
    AddReportLog "RemarkingBuyOut_Click"
    DoCmd.close acForm, "RemarketingBuyOut"
    DoCmd.OpenForm "RemarketingBuyOut", acNormal
End Sub

Private Sub RemarketingInvReport_Click()
    If Me.AsOfDate & "test" = "test" Then
        MsgBox "You must have an As-Of date", vbInformation, "Date Req"
        Exit Sub
    End If
    '' Update RMKT Data if need be
    If DateDiff("h", DLookup("LastWriteDate", "LastWriteReport"), Now()) >= 3 Then
        MessageUser "Open", "Updating from Server", "This will be a moment..."
        DoCmd.SetWarnings False
        DoCmd.RunSQL ("SELECT #" & Now() & "# As LastWriteDate INTO [LastWriteReport];")
        Call SetRemarketingAllUnits
        DoCmd.SetWarnings True
        MessageUser "Close"
    End If
    Dim XLS, DateX As String
    ' Dim sCount As Integer
    DateX = Format(Me.AsOfDate, "MMM-DD-YYYY")

    SQLz = "SELECT "
    SQLz = SQLz & "SourceCompany as Client"                        'Col A
    SQLz = SQLz & ",UnitId as AssetID"                             'Col B
    SQLz = SQLz & ",UnitNum"                                       'Col C
    SQLz = SQLz & ",UnitVIN As VIN"                                'Col D
    SQLz = SQLz & ",UnitDescShortVer as [Description]"             'Col E
    SQLz = SQLz & ",ParkedLocation as [Parked Location]"           'Col F
    SQLz = SQLz & ",ParkedAddress & ' ' & PCity & ', ' & PState & '. ' & PZip as [Parked Address]"             'Col G
    SQLz = SQLz & ",PState AS [Park St]"                           'Col H
    SQLz = SQLz & ",BuyOut AS [BuyOut Cost]"                       'Col I
   '' Sqlz = Sqlz & ",(BuyOut + nz(TermPenalty,0)) as [Final Cost]"  'Col J
    SQLz = SQLz & ",BuyOutDate As [Buy Out Date]"                  'Col J
    SQLz = SQLz & ",InvoicePaid as [Invoice Paid Date]"            'Col K
    SQLz = SQLz & ",UnitStatus AS [Status]"                        'Col L
    SQLz = SQLz & ",IFBank As [Interim Bank]"                      'Col M
    SQLz = SQLz & ",IFAmount AS [Interim Amount]"                  'Col N
    SQLz = SQLz & ",iif(IFAmount>0, (BuyOut - IFAmount),NULL) AS [Interim Var]" 'Col O
    SQLz = SQLz & ",IFLoanNumber AS AccountNumber"                 'Col P
    SQLz = SQLz & ",IFDate As InterimDate"                         'Col Q
    SQLz = SQLz & ",IFPaidDate AS [Interim Paid Date]"             'Col R
    SQLz = SQLz & ",TermPenalty AS [Term penalty]"                 'Col S
    SQLz = SQLz & ",BuyOutPct AS [Buy Out Pct]"                    'Col T
    SQLz = SQLz & ",Buyer"                                         'Col U
    SQLz = SQLz & ",UnitSalePrice as [Actual Sold Price]"          'Col V
    SQLz = SQLz & ",OffLeaseDate AS [Actual OffLease]"             'Col W
    SQLz = SQLz & ",Disposition as DSPN"                           'Col X
    SQLz = SQLz & ",InspectionPassed as [Insp Passed]"             'Col Y
    SQLz = SQLz & " FROM (Select * From RemarketingReportAllUnits WHERE NOT UnitStatus in ('SB','C','N')) as RR"
    SQLz = SQLz & " WHERE (((InvoicePaid) Is Null Or (InvoicePaid) > #" & AsOfDate & "#) AND ((BuyOutDate) <=#" & AsOfDate & "#))"

  '  Sqlz = "SELECT AR.* FROM AcctRMKTInventoryRpt as AR;"

    'sCount = DCount("AssetId", "AcctRMKTInventoryRpt")
    MessageUser "Open", "Getting RMKT Inventory", "This will be a moment..."
    AddReportLog "RemarketingInvReport"
    XLS = "|A,DD3|A,mTable:Name=Inventory|B,BC|B,TX|C,TC|E,TL|D,TC|H,BC|I,$E|J,TC|K,TC|L,BC|L,TX|M,TC|N,$E|O,$E|P,TC|Q,TC|R,TC|S,$E|T,P1|T,TC|V,$E|W,TC|Y,TC|"
    XLS = XLS & "B,cBorder:Med|I,cBorder:Med|L,cBorder:Med|N,cBorder:Med|S,cBorder:Med|V,cBorder:Med|B,W09|H,W06|V,W13|I,W13|J,W13|K,W13|N,W14|Q,W13|R,W13|S,W14|V,W14|"
    ''Z,$E|AA,$E|AB,$E|AC,$R|AD,$R|AE,$R|AF,$R|"
    SendReport2Excel SQLz, "Inventory", "RMKT Inventory: " & DateX, True, XLS, 0
    ''"C:\Faas\" & FAUserName & "\Temp\RMKTInventory " & DateX & ".xlsx"
    'FormatExcel "AcctRMKTInventoryRpt", "C:\Faas\" & FAUserName & "\Temp\RMKTInventory " & DateX & ".xlsx", "RMKT Inventory", "", XLS, 0, "Remarketing Inventory: " & DateX, True
    MessageUser "Close"
End Sub

Private Sub ReserveReport_Click()
    If Me.StartDate & "test" = "test" Or Me.EndDate & "test" = "test" Then
        MsgBox "You must have an BOTH a Start and End Date", vbInformation, "Date Req"
        Exit Sub
    End If
    Dim XLS, DateX As String
    Dim sCount As Integer
    DateX = Format(Me.StartDate, "MMM-DD-YYYY") & "~" & Format(Me.EndDate, "MMM-DD-YYYY")
    'sCount = DCount("AssetId", "AcctRMKTInventoryReserveRpt")
    XLS = "|B,BC|B,TX|C,TC|E,BL|D,TC|H,$E|K,BC|L,BL|M,$E|N,BL|Q,$E|R,P1R|T,$E|U,$E|V,$E|"
    AddReportLog "AcctRMKTInventoryReserveRpt"
    FormatExcel "AcctRMKTInventoryReserveRpt", "C:\Faas\" & FAUserName & "\Temp\RMKTReserveReport " & DateX & ".xlsx", "RMKT Reserve-Profit", "", XLS, sCount, "Remarketing Reserve vs. Profit: " & DateX, True
End Sub

Private Sub SelClient_Click()
    Me.SelLease = ""
    Me.SelSch = 0
    Me.SelGrp = 0
    Me.SelLease.Requery
    Me.SelSch.Requery
    Me.SelGrp.Requery
End Sub

Private Sub SelLease_Click()
    Me.SelSch = 0
    Me.SelGrp = 0
    Me.SelSch.Requery
    Me.SelGrp.Requery
End Sub

Private Sub SelSch_Click()
    Me.SelGrp = 0
    Me.SelGrp.Requery
End Sub

Private Sub SurrenderPDReport_Click()
    Dim SQLz, XLS As String
    Call MessageUser("Open", "Getting Data...", "Excel will open in a bit.")
    SQLz = "Select * From vw_PotentialBillingSurrPerDiem "
    SQLz = SQLz & " ORDER BY Client,SchNo,GroupNo"

    AddReportLog "SurrenderPDReport"
    XLS = "|A,W15|E,TC|F,BC|F,TX|I,TC|R,BC|T,$E|U,$E|U,$E|V,TC|W,P3|X,$E|Y,$E|Z,$R|AA,TC|"
    Call SendReport2Excel(SQLz, "Billable", "Surrender Per-Diem Report:" & Date, True, XLS, 0, True)
    Call MessageUser("Close")
End Sub

Private Sub TESRptBtn_Click()
    'Call OpenTESRPTPDF
End Sub

Private Sub UlookBtn_Click()
    DoCmd.OpenForm "UnitReverseLookUpForm"
End Sub
