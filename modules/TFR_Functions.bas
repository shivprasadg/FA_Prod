Attribute VB_Name = "TFR_Functions"
Option Compare Database
Option Explicit
Public Function MakeTFRReport(xType As String, xReportName As String, Optional StartDateRange As String, Optional EndDateRange As String, _
                              Optional ClientGroupIDi As Integer, Optional BankIDi As Integer, Optional SaveToPDF As Boolean)
    If FAUserAccessGroup Like "*ACCT*" Or FAUserAccessGroup Like "*TMT*" Then
        Call MakeTFRWorkTable
        Dim SQLz As String
        Dim qdf As DAO.QueryDef
        Dim pFileName, ReportType As String
        Dim SqlDate, SqlClient, SqlBank, SqlNull, DateRangeX As String
        If StartDateRange = "" Then StartDateRange = "1/1/" & DatePart("yyyy", Date)
        If EndDateRange = "" Then EndDateRange = "12/31/" & DatePart("yyyy", Date)

        pFileName = "C:\temp\TFR_" & DatePart("yyyy", Date) & "_TFRDetail_" & Format(Date, "mm_dd_yyyy")
        pFileName = pFileName & ".pdf"
        Call MessageUser("Open", xReportName, "This will take just a moment or two...")
        SQLz = "Select * From TFR_Work WHERE Assigned_Indicator='I' OR Assigned_Indicator <> '' "
        SqlDate = "AND FundDate >=#" & StartDateRange & "# "
        SqlDate = SqlDate & "AND FundDate <=#" & EndDateRange & "# "
        DateRangeX = StartDateRange & " ~ " & EndDateRange
        If Nz(ClientGroupIDi, 0) <> 0 Then SqlClient = "AND ClientGroupID=" & ClientGroupIDi & " "
        If Nz(BankIDi, 0) <> 0 Then SqlBank = "AND BankId=" & BankIDi & " "
        SqlNull = "Or FundDate Is Null "
        SqlDate = SqlDate & SqlClient & SqlBank
        SqlNull = SqlNull & SqlClient & SqlBank
        SQLz = SQLz & SqlDate & SqlNull
        On Error Resume Next
        DoCmd.DeleteObject acQuery, "TFRCompare"
        On Error GoTo 0
        Set qdf = CurrentDb.CreateQueryDef("TFRCompare", SQLz)
        If xType = "Report" Then
           DoCmd.OpenReport xReportName, acViewReport, , , , DateRangeX
           Else
           DoCmd.OpenReport xReportName, acViewPreview, , , , DateRangeX
           End If
        If SaveToPDF = True Then DoCmd.OutputTo acOutputReport, xReportName & ", acFormatPDF, pFileName,True"
        Call MessageUser("Close")
    Else
        MsgBox "Sorry you do NOT have access to this report.", vbCritical, AtlaasVersion

    End If
End Function
Public Function MakeTFRWorkTable(Optional Override As Boolean)
    If Round(Now() - CDate(DMax("DataDate", "TFR_Work")), 1) > 0.29 Or Override = True Then
        If IsOpen("TFRDashboard") Then
            Dim FormRecX, AccountRecX, FundYearX, BankTypeX, DescRecX, Sch_GrpRecX, StatusRecX, BDEXRecX, CGroupRecX As String
            FormRecX = Forms!TFRDashboard.Form.RecordSource
            AccountRecX = Forms!TFRDashboard.lstAccountName.RowSource
            FundYearX = Forms!TFRDashboard.lstFundYear.RowSource
            BankTypeX = Forms!TFRDashboard.lstBankType.RowSource
            DescRecX = Forms!TFRDashboard.lstDesc.RowSource
            Sch_GrpRecX = Forms!TFRDashboard.lstSch_Grp.RowSource
            StatusRecX = Forms!TFRDashboard.lstStatus.RowSource
            BDEXRecX = Forms!TFRDashboard.lstBDEX.RowSource
            CGroupRecX = Forms!TFRDashboard.lstCGroup.RowSource

            Forms!TFRDashboard.Form.RecordSource = ""
            Forms!TFRDashboard.lstAccountName.RowSource = ""
            Forms!TFRDashboard.lstFundYear.RowSource = ""
            Forms!TFRDashboard.lstBankType.RowSource = ""
            Forms!TFRDashboard.lstDesc.RowSource = ""
            Forms!TFRDashboard.lstSch_Grp.RowSource = ""
            Forms!TFRDashboard.lstStatus.RowSource = ""
            Forms!TFRDashboard.lstBDEX.RowSource = ""
            Forms!TFRDashboard.lstCGroup.RowSource = ""
        End If
        DoCmd.close acReport, "TFRCompareSummary"
        DoCmd.close acReport, "TFRCompare"

        DoCmd.SetWarnings False
        DoCmd.RunSQL ("SELECT Now() as DataDate, *  INTO TFR_Work FROM TFR_2;")
        DoCmd.SetWarnings True

        If IsOpen("TFRDashboard") Then
            Forms!TFRDashboard.lstAccountName.RowSource = AccountRecX
            Forms!TFRDashboard.lstFundYear.RowSource = FundYearX
            Forms!TFRDashboard.lstBankType.RowSource = BankTypeX
            Forms!TFRDashboard.lstDesc.RowSource = DescRecX
            Forms!TFRDashboard.lstSch_Grp.RowSource = Sch_GrpRecX
            Forms!TFRDashboard.lstStatus.RowSource = StatusRecX
            Forms!TFRDashboard.lstBDEX.RowSource = BDEXRecX
            Forms!TFRDashboard.lstCGroup.RowSource = CGroupRecX
            Forms!TFRDashboard.Form.RecordSource = FormRecX
        End If
    End If
End Function
Public Function QuickTFR2Excel(SQLg As String)
Call MessageUser("Open", "Quick TRF Report", "This will be quick...")
Dim XLS, SQLx As String
    SQLx = SQLx & "Select "
    SQLx = SQLx & " ClientGroup"
    SQLx = SQLx & ",AccountName"
    SQLx = SQLx & ",Sch_Grp"
    SQLx = SQLx & ",UnitDesc"
    SQLx = SQLx & ",UnitCount"
    SQLx = SQLx & ",LPAFSigned"
    SQLx = SQLx & ",FundDate"
    SQLx = SQLx & ",OECPerUnit"
    SQLx = SQLx & ",OEC"
    SQLx = SQLx & ",Bank"
    SQLx = SQLx & ",SyndFee"
    SQLx = SQLx & ",SyndFeePct"
    SQLx = SQLx & ",TermPenalty"
    SQLx = SQLx & ",TermPenaltyPct"
    SQLx = SQLx & ",InterimExp"
    SQLx = SQLx & ",NetSyndFee"
    SQLx = SQLx & ",NetSyndPct"
    SQLx = SQLx & ",PDRent"
    SQLx = SQLx & ",PDRentPct"
    SQLx = SQLx & ",PDDaysAvg"
    SQLx = SQLx & ",UTA"
    SQLx = SQLx & ",UTApct"
    SQLx = SQLx & ",NPV"
    SQLx = SQLx & ",NPVPct"
    SQLx = SQLx & ",TotalFee"
    SQLx = SQLx & ",TotalFeePct"
    SQLx = SQLx & ",PmtAmt"
    SQLx = SQLx & ",FMSS"
    SQLx = SQLx & ",PmtAmtFMSS"
    SQLx = SQLx & ",RVPerUnit"
    SQLx = SQLx & ",RVPct"
    SQLx = SQLx & ",LeaseTermPrimary"
    SQLx = SQLx & ",MilesAllowed"
    SQLx = SQLx & ",SWAP"
    SQLx = SQLx & ",Spread"
    SQLx = SQLx & ",Yield"
    SQLx = SQLx & ",Bonus"
    SQLx = SQLx & ",DealCommission "

SQLx = SQLx & " From" & SplitString(SQLg, " From", 1)
XLS = "|A,DD3|"
    XLS = XLS & "A,TL|A,cLabel:End=A:3:Client" & vbNewLine & "Group|"
    XLS = XLS & "B,TL|B,cLabel:End=B:3:Account" & vbNewLine & "Name|"
    XLS = XLS & "C,TL|C,cLabel:End=C:3:Schedule" & vbNewLine & "Group|C,BX|"
    XLS = XLS & "D,TL|D,cLabel:End=D:3:Unit" & vbNewLine & "Desc.|"
    XLS = XLS & "E,TC|E,cLabel:End=E:3:Unit" & vbNewLine & "Count|E,BE|"
    XLS = XLS & "F,TC|F,cLabel:End=F:3:Awarded" & vbNewLine & "Date|"
    XLS = XLS & "G,TC|G,cLabel:End=G:3:Fund" & vbNewLine & "Mth" & vbNewLine & "Year|G,MthYear|"
    XLS = XLS & "H,$IntE|H,cLabel:End=H:3:OEC" & vbNewLine & "Unit" & vbNewLine & "(avg)|"
    XLS = XLS & "I,$IntE|I,cLabel:End=I:3:Group" & vbNewLine & "OEC|"
    XLS = XLS & "J,BL|"
    XLS = XLS & "K,$IntE|K,cLabel:End=K:3:Gross" & vbNewLine & "Synd" & vbNewLine & "Fee|"
    XLS = XLS & "L,P2|L,cLabel:End=L:3:Gross" & vbNewLine & "Synd" & vbNewLine & "%|L,FColor:47.117.181|"
    XLS = XLS & "M,$IntE|M,cLabel:End=M:3:Term" & vbNewLine & "Penalty|"
    XLS = XLS & "N,P2|N,cLabel:End=N:3:Term" & vbNewLine & "Penalty" & vbNewLine & "%|N,FColor:47.117.181|"
    XLS = XLS & "O,$IntE|O,cLabel:End=O:3:Interim" & vbNewLine & "Exp|"
    XLS = XLS & "P,$IntE|P,cLabel:End=P:3:Net" & vbNewLine & "Synd" & vbNewLine & "Fee|"
    XLS = XLS & "Q,P2|Q,cLabel:End=Q:3:Net" & vbNewLine & "Synd" & vbNewLine & "%|Q,FColor:47.117.181|"
    XLS = XLS & "R,$IntE|R,cLabel:End=R:3:PerDiem" & vbNewLine & "Group|"
    XLS = XLS & "S,P2|S,cLabel:End=S:3:Per" & vbNewLine & "Diem" & vbNewLine & "%|S,FColor:47.117.181|"
    XLS = XLS & "T,BC|T,cLabel:End=T:3:PerDiem" & vbNewLine & "Days" & vbNewLine & "(avg)|"
    XLS = XLS & "U,$IntE|"
    XLS = XLS & "V,P2|V,cLabel:End=V:3:UTA" & vbNewLine & "%|V,FColor:47.117.181|"
    XLS = XLS & "W,$IntE|W,cLabel:End=W:3:FMSS" & vbNewLine & "NPV|"
    XLS = XLS & "X,P2|X,cLabel:End=X:3:NPV" & vbNewLine & "%|X,FColor:47.117.181|"
    XLS = XLS & "Y,$IntE|Y,cLabel:End=Y:3:Total" & vbNewLine & "Fee|"
    XLS = XLS & "Z,P2|Z,cLabel:End=Z:3:Total" & vbNewLine & "Fee" & vbNewLine & "%|Z,FColor:47.117.181|"
    XLS = XLS & "AA,$IntE|AA,cLabel:End=AA:3:Payment" & vbNewLine & "Amt|"
    XLS = XLS & "AB,$IntE|"
    XLS = XLS & "AC,$IntE|AC,cLabel:End=AC:3:PMT" & vbNewLine & "FMSS|"
    XLS = XLS & "AD,$IntE|AD,cLabel:End=AD:3:RV" & vbNewLine & "Unit" & vbNewLine & "avg|"
    XLS = XLS & "AE,P2|AE,cLabel:End=AE:3:RV" & vbNewLine & "Unit" & vbNewLine & "%|AE,FColor:0.128.0|"
    XLS = XLS & "AF,BC|AF,cLabel:End=AF:3:Term" & vbNewLine & "Mths|"
    XLS = XLS & "AG,SR|AG,cLabel:End=AG:3:Miles" & vbNewLine & "PYear|"
    XLS = XLS & "AH,P2C|"
    XLS = XLS & "AI,P2C|"
    XLS = XLS & "AJ,P2C|"
    XLS = XLS & "AK,%C|AK,cLabel:End=AK:3:Bonus" & vbNewLine & "Depr|"
    XLS = XLS & "AL,P2C|AL,cLabel:End=AL:3:Deal" & vbNewLine & "Comm|"
    XLS = XLS & "A,xlCenter|"
    XLS = XLS & "B,LLT|B,TL|C,W15|D,W16|E,W07|J,W08|L,W07|N,W07|O,W09|Q,W08|S,W07|T,W07|V,W07|X,W07|Z,W08|AA,W10|"
    XLS = XLS & "AE,W07|AF,W06|AG,W07|AH,W07|AI,W07|AJ,W07|AK,W07|AL,W06|"
SendReport2Excel SQLx, "QuickTFR ", "Quick TFR" & Format(Date, "mm-dd-yy"), True, XLS, 0, True
Call MessageUser("Close")
End Function
Public Function MakeTFRShort(Optional StartDateRange As String, Optional EndDateRange As String, Optional ClientGroupIDi As Integer, Optional BankIDi As Integer)

    If FAUserAccessGroup Like "*ACCT*" Or FAUserAccessGroup Like "*ITT*" Or FAUserAccessGroup Like "*TMT*" Then
        'Call MakeTFRWorkTable
    Else
        MsgBox "Sorry you do NOT have access to this report.", vbCritical, AtlaasVersion
        Exit Function
    End If
    Dim SQLz As String
    Dim SQLy As String
    Dim SQLx As String
    Dim SqlA As String
    Dim Sqlb As String
    Dim SqlDate, SqlClient, SqlBank, SqlNull, DateRangeX As String
    Dim Go1, Go2, Go3, Go4, NumRpts, GrpRpt As Integer
    Dim IsLastRpt As Boolean
    Dim XLS, TitleX, TitleDateX As String
    Call MessageUser("Open", "Setting up TFR", "This will take a few moments")

    SQLx = "Select ClientGroup" '- Col:A
    SQLx = SQLx & " ,AccountName " '- Col:B
    SQLx = SQLx & " ,Sch_Grp" '- Col:C
    SQLx = SQLx & " ,FAID" '- Col:D
    SQLx = SQLx & " ,LPAFYear" '- Col:E
    SQLx = SQLx & " ,Assigned_Indicator" '- Col:F
    SQLx = SQLx & " ,FundDate" '- Col:G
    SQLx = SQLx & " ,Bank" '- Col:H
    SQLx = SQLx & " ,LeaseTermPrimary" '- Col:I
    SQLx = SQLx & " ,UnitCount" '- Col:J
    SQLx = SQLx & " ,DescAbbrv" '- Col:K
    SQLx = SQLx & " ,OECPerUnit" '- Col:L
    SQLx = SQLx & " ,GroupOEC" '- Col:M
    SQLx = SQLx & " ,VendorCostUnit" '- Col:N
    SQLx = SQLx & " ,TotalVendorCost" '- Col:O
    SQLx = SQLx & " ,SyndFee" '- Col:P
    SQLx = SQLx & " ,NetSyndFee" '- Col:Q
    SQLx = SQLx & " ,TotalFee" '- Col:R
    SQLx = SQLx & " ,SWAPType" '- Col:S
    SQLx = SQLx & " ,Yield" '- Col:T
    SQLx = SQLx & " ,Spread" '- Col:U
    SQLx = SQLx & " ,IndexRate" '- Col:V
    SQLx = SQLx & " ,SpreadIRR" '- Col:W
    SQLx = SQLx & " ,SyndFeeProj" '- Col:X
    SQLx = SQLx & " ,SpreadProj" '- Col:Y
    SQLx = SQLx & " ,LoAmount" '- Col:Z
    SQLx = SQLx & " ,(LoAmount * UnitCount) As LOGroupOEM" '- Col:AA
    SQLx = SQLx & " ,BankContract" '- Col:AB
    SQLx = SQLx & " From TFR_2 "


    XLS = "|"
    XLS = XLS & "A,DD|A,TR|A,cLabel:End=A:3:Client " & vbNewLine & "Group|"
    XLS = XLS & "B,BL|B,cLabel:End=B:3:Account" & vbNewLine & "Name|"
    XLS = XLS & "C,cLabel:End=C:3:Schedule" & vbNewLine & "Group|"
    XLS = XLS & "D,TL|D,cLabel:End=D:3:FAID|"
    XLS = XLS & "E,BC|E,cLabel:End=E:3:LPAF" & vbNewLine & "Fiscal" & vbNewLine & "Year|E,cColor:3:0.0.0|"
    XLS = XLS & "F,BC|F,cLabel:End=F:3:Fund" & vbNewLine & "Status|F,cColor:3:0.0.0|"
    XLS = XLS & "G,cLabel:End=G:3:Fund" & vbNewLine & "Date|G,cColor:3:0.0.0|"
    XLS = XLS & "H,BL|H,cLabel:End=H:3:Bank|"
    XLS = XLS & "I,BC|I,cLabel:End=I:3:Term" & vbNewLine & "Mths|"
    XLS = XLS & "J,BC|J,cLabel:End=J:3:Unit" & vbNewLine & "Count|"
    XLS = XLS & "K,TL|K,cLabel:End=K:3:Desc|"
    XLS = XLS & "L,$IntE|L,cLabel:End=L:3:OEC" & vbNewLine & "Unit" & vbNewLine & "(avg)|"
    XLS = XLS & "M,$IntE|M,cLabel:End=M:3:Group" & vbNewLine & "OEC|M,cColor:3:0.0.50|"
    XLS = XLS & "N,$IntE|N,cLabel:End=N:3:Vendor" & vbNewLine & "Unit" & vbNewLine & "(avg)|"
    XLS = XLS & "O,$IntE|O,cLabel:End=O:3:Group" & vbNewLine & "Vendor" & vbNewLine & "Cost|O,cColor:3:0.0.0|"
    XLS = XLS & "P,$IntE|P,cLabel:End=P:3:Gross" & vbNewLine & "Synd" & vbNewLine & "Fee|"
    XLS = XLS & "Q,$IntE|Q,cLabel:End=Q:3:Net" & vbNewLine & "Synd" & vbNewLine & "Fee|"
    XLS = XLS & "R,$IntE|R,cLabel:End=R:3:Total" & vbNewLine & "Fee|"
    XLS = XLS & "S,BL|S,cLabel:End=S:3:SWAP" & vbNewLine & "Type|S,cColor:3:0.0.0|"
    XLS = XLS & "T,P2|T,TC|T,cLabel:End=T:3:Yield|T,cColor:3:55.86.35|"
    XLS = XLS & "U,P2|U,TC|U,cLabel:End=U:3:Spread|U,cColor:3:55.86.35|"
    XLS = XLS & "V,P2|V,TC|U,cLabel:End=V:3:Index" & vbNewLine & "Rate|V,cColor:3:191.143.0|"
    XLS = XLS & "W,P2|W,TC|W,cLabel:End=W:3:Spread" & vbNewLine & "IRR-PT|W,cColor:3:84.130.53|"
    XLS = XLS & "X,$IntE|X,cLabel:End=X:3:LO Synd" & vbNewLine & "Fee|X,cColor:3:150.150.150|"
    XLS = XLS & "Y,P2|X,TC|Y,cLabel:End=Y:3:LO " & vbNewLine & "Spread|Y,cColor:3:150.150.150|"
    XLS = XLS & "Z,$IntE|Z,cLabel:End=Y:Z:LO OEC" & vbNewLine & "Unit" & vbNewLine & "(avg)|Z,cColor:3:120.120.120|"
    XLS = XLS & "AA,$IntE|AA,cLabel:End=AA:3:LO Group" & vbNewLine & "OEC|AA,cColor:3:100.100.100|"
    XLS = XLS & "A,xlCenter|"
    XLS = XLS & "B,LLT|B,TL|"

    TitleX = ""
    TitleDateX = ""
    If IsDate(StartDateRange) Then TitleDateX = TitleDateX & StartDateRange & " "
    If IsDate(StartDateRange) And IsDate(EndDateRange) Then TitleDateX = TitleDateX & "~ "
    If IsDate(EndDateRange) Then TitleDateX = TitleDateX & EndDateRange & " "
    If Nz(ClientGroupIDi, 0) <> 0 Then TitleX = TitleX & "Client " & ClientGroupIDi & " "
    If Nz(BankIDi, 0) <> 0 Then TitleX = TitleX & "Bank " & DLookup("BankShortName", "Banks", "BankId=" & BankIDi) & " "

    If IsDate(StartDateRange) Then SqlDate = "AND FundDate >=#" & StartDateRange & "# "
    If IsDate(EndDateRange) Then SqlDate = SqlDate & "AND FundDate <=#" & EndDateRange & "# "
    If Nz(ClientGroupIDi, 0) <> 0 Then SqlClient = "AND ClientGroupID=" & ClientGroupIDi & " "
    If Nz(BankIDi, 0) <> 0 Then SqlBank = "AND BankId=" & BankIDi & " "
    If IsDate(EndDateRange) Or IsDate(StartDateRange) Then SqlNull = "Or FundDate Is Null "
    SqlDate = SqlDate & SqlClient & SqlBank
    SqlNull = SqlNull & SqlClient & SqlBank
    SQLx = SQLx & " WHERE Assigned_Indicator = 'F' " & SqlDate
    Call MessageUser("Update", "Short TRF For All Groups", "Setting Funded, just a sec...")
    Call SendReport2Excel(SQLx, "ShortTFR", "Short TFR " & TitleX & TitleDateX, True, XLS, 0, 0, 0, , , "C:\Faas\" & FaUserName & "\Temp\ShortTFRReport.xlsx")
    Call MessageUser("Close")
End Function
Public Function MakeTMTTFR(Optional StartDateRange As String, Optional EndDateRange As String, Optional ClientGroupIDi As Integer, Optional BankIDi As Integer, Optional FundStatus As Integer)

    If FAUserAccessGroup Like "*ACCT*" Or FAUserAccessGroup Like "*ITT*" Or FAUserAccessGroup Like "*TMT*" Then
        'Call MakeTFRWorkTable
    Else
        MsgBox "Sorry you do NOT have access to this report.", vbCritical, AtlaasVersion
        Exit Function
    End If
    Dim SQLz As String
    Dim SQLy As String
    Dim SQLx As String
    Dim SqlA As String
    Dim Sqlb As String
    Dim SqlDate, SqlClient, SqlBank, SqlNull, DateRangeX As String
    Dim Go1, Go2, Go3, Go4, NumRpts, GrpRpt As Integer
    Dim IsLastRpt As Boolean
    Dim XLS, TitleX, TitleDateX As String
    Call MessageUser("Open", "Setting up TMT-TFR", "This will take a few moments")

    SQLz = "Select AccountName" '- Col:A
    SQLz = SQLz & " ,FAID" '- Col:B
    SQLz = SQLz & " ,Sch_Grp" '- Col:C
    SQLz = SQLz & " ,FundDate" '- Col:D
    SQLz = SQLz & " ,Bank" '- Col:E
    SQLz = SQLz & " ,Assigned_Indicator" '- Col:F
    SQLz = SQLz & " ,LeaseType" '- Col:G
    SQLz = SQLz & " ,LeaseTermPrimary" '- Col:H
    SQLz = SQLz & " ,MilesAllowed" '- Col:I
    SQLz = SQLz & " ,HoursAllowed" '- Col:J
    SQLz = SQLz & " ,UnitCount" '- Col:K
    SQLz = SQLz & " ,Myear" '- Col:L
    SQLz = SQLz & " ,Make" '- Col:M
    SQLz = SQLz & " ,Model" '- Col:N
    SQLz = SQLz & " ,Type" '- Col:O
    SQLz = SQLz & " ,SubType" '- Col:P
    SQLz = SQLz & " ,Axle" '- Col:Q
    SQLz = SQLz & " ,EngineType" '- Col:R
    SQLz = SQLz & " ,TransType" '- Col:S
    SQLz = SQLz & " ,LOUnitCost" '- Col:T
    SQLz = SQLz & " ,LOCostGroup" '- Col:U
    SQLz = SQLz & " ,LORV" '- Col:V
    SQLz = SQLz & " ,LORVGroup" '- Col:W
    SQLz = SQLz & " ,LORVPct" '- Col:X
    SQLz = SQLz & " ,LORent" '- Col:Y
    SQLz = SQLz & " ,LORentGroup" '- Col:Z
    SQLz = SQLz & " ,SchLRF" '- Col:AA
    SQLz = SQLz & " ,LOSwapType" '- Col:AB
    SQLz = SQLz & " ,LOSwap" '- Col:AC
    SQLz = SQLz & " ,LOSwapGroup" '- Col:AD
    SQLz = SQLz & " ,LOSpread" '- Col:AE
    SQLz = SQLz & " ,LOSpreadGroup" '- Col:AF
    SQLz = SQLz & " ,LOYield" '- Col:AG
    SQLz = SQLz & " ,LOYieldGroup" '- Col:AH
    SQLz = SQLz & " ,LOBonus" '- Col:AI
    SQLz = SQLz & " ,LOBonusGroup" '- Col:AJ
    SQLz = SQLz & " ,LOSyndFee" '- Col:AK
    SQLz = SQLz & " ,LOSyndFeeGroup" '- Col:AL
    SQLz = SQLz & " ,LoSyndPct" '- Col:AM
    SQLz = SQLz & " ,LOTermPenalty" '- Col:AN
    SQLz = SQLz & " ,LOTermPenaltyGroup" '- Col:AO
    SQLz = SQLz & " ,LoTermPenaltyPct" '- Col:AP
    SQLz = SQLz & " ,LoSyndNet" '- Col:AQ
    SQLz = SQLz & " ,LoSyndNetGroup" '- Col:AR
    SQLz = SQLz & " ,LoSyndNetPct" '- Col:AS
    SQLz = SQLz & " ,LOPDUnit" '- Col:AT
    SQLz = SQLz & " ,LOPDUnitGroup" '- Col:AU
    SQLz = SQLz & " ,LoPDPct" '- Col:AV
    SQLz = SQLz & " ,LoUTA" '- Col:AW
    SQLz = SQLz & " ,LoUTAGroup" '- Col:AX
    SQLz = SQLz & " ,LoUTAPct" '- Col:AY
    SQLz = SQLz & " ,LOTotalFee" '- Col:AZ
    SQLz = SQLz & " ,LOTotalFeeGroup" '- Col:BA
    SQLz = SQLz & " ,LoTotalFeePct" '- Col:BB
    SQLz = SQLz & " ,UnitCost" '- Col:BC
    SQLz = SQLz & " ,UnitCostGroup" '- Col:BD
    SQLz = SQLz & " ,PartsCost" '- Col:BE
    SQLz = SQLz & " ,PartsCostGroup" '- Col:BF
    SQLz = SQLz & " ,TotalVendorCost" '- Col:BG
    SQLz = SQLz & " ,TotalVendorCostGroup" '- Col:BH
    SQLz = SQLz & " ,RVPerUnit" '- Col:BI
    SQLz = SQLz & " ,RVPerUnitGroup" '- Col:BJ
    SQLz = SQLz & " ,RVVendorPct" '- Col:BK
    SQLz = SQLz & " ,RVPct" '- Col:BL
    SQLz = SQLz & " ,PmtAmt" '- Col:BM
    SQLz = SQLz & " ,PmtAmtGroup" '- Col:BN
    SQLz = SQLz & " ,LRF" '- Col:BO
    SQLz = SQLz & " ,SWAPType" '- Col:BP
    SQLz = SQLz & " ,SWAP" '- Col:BQ
    SQLz = SQLz & " ,SWAPGroup" '- Col:BR
    SQLz = SQLz & " ,Spread" '- Col:BS
    SQLz = SQLz & " ,SpreadGroup" '- Col:BT
    SQLz = SQLz & " ,Yield" '- Col:BU
    SQLz = SQLz & " ,YieldGroup" '- Col:BV
    SQLz = SQLz & " ,Bonus" '- Col:BW
    SQLz = SQLz & " ,BonusGroup" '- Col:BX
    SQLz = SQLz & " ,SyndFee" '- Col:BY
    SQLz = SQLz & " ,SyndFeeGroup" '- Col:BZ
    SQLz = SQLz & " ,SyndFeePct" '- Col:CA
    SQLz = SQLz & " ,TermPenalty" '- Col:CB
    SQLz = SQLz & " ,TermPenaltyGroup" '- Col:CC
    SQLz = SQLz & " ,TermPenaltyPct" '- Col:CD
    SQLz = SQLz & " ,NetSyndFee" '- Col:CE
    SQLz = SQLz & " ,NetSyndFeeGroup" '- Col:CF
    SQLz = SQLz & " ,NetSyndPct" '- Col:CG
    SQLz = SQLz & " ,PDRent" '- Col:CH
    SQLz = SQLz & " ,PDRentGroup" '- Col:CI
    SQLz = SQLz & " ,PDRentPct" '- Col:CJ
    SQLz = SQLz & " ,UTA" '- Col:CK
    SQLz = SQLz & " ,UTAGroup" '- Col:CL
    SQLz = SQLz & " ,UTAPct" '- Col:CM
    SQLz = SQLz & " ,TotalFee" '- Col:CN
    SQLz = SQLz & " ,TotalFeeGroup" '- Col:CO
    SQLz = SQLz & " ,TotalFeePct" '- Col:CP
    SQLz = SQLz & " From vw_TFR_TMT2 "


    XLS = "|A,DD4|"
    XLS = XLS & "A,TC|A,cLabel:End=A:3:Client|A,cColor:3:0.0.0|A,W14|"
    XLS = XLS & "B,TC|B,cLabel:End=B:3:FAID|B,cColor:3:0.0.0|B,W12|"
    XLS = XLS & "C,TC|C,cLabel:End=C:3:Schedule Group|C,cColor:3:0.0.0|C,W10|"
    XLS = XLS & "D,MthYear|D,cLabel:End=D:3:Fund Date|D,cColor:3:0.0.0|D,TC|D,W10|"
    XLS = XLS & "E,TC|E,cLabel:End=E:3:Bank|E,cColor:3:0.0.0|E,W07|"
    XLS = XLS & "F,TC|F,cLabel:End=F:3:Fund Status|F,cColor:3:0.0.0|F,W05|"
    XLS = XLS & "G,TC|G,cLabel:End=G:3:Lease Type|G,cColor:3:0.0.0|G,W11|"
    XLS = XLS & "H,AC|H,cLabel:End=H:3:Term|H,cColor:3:0.0.0|H,TC|H,W06|"
    XLS = XLS & "I,Sint|I,AC|I,cLabel:End=I:3:MPY|I,cColor:3:0.0.0|I,TC|I,W08|"
    XLS = XLS & "J,Sint|J,AC|J,cLabel:End=J:3:HPY|J,cColor:3:0.0.0|J,TC|J,W08|"
    XLS = XLS & "K,SE|K,cLabel:End=K:3:Unit Count|K,cColor:3:0.0.0|K,TC|K,W08|"
    XLS = XLS & "L,TC|L,cLabel:End=L:3:Model Year|L,cColor:3:0.0.0|L,W06|"
    XLS = XLS & "M,TC|M,cLabel:End=M:3:Make|M,cColor:3:0.0.0|M,W12|"
    XLS = XLS & "N,TC|N,cLabel:End=N:3:Model|N,cColor:3:0.0.0|N,W11|"
    XLS = XLS & "O,TC|O,cLabel:End=O:3:Type|O,cColor:3:0.0.0|O,W12|"
    XLS = XLS & "P,TC|P,cLabel:End=P:3:Sub Type|P,cColor:3:0.0.0|P,W12|"
    XLS = XLS & "Q,TC|Q,cLabel:End=Q:3:Axle|Q,cColor:3:0.0.0|Q,W05|"
    XLS = XLS & "R,TC|R,cLabel:End=R:3:Engine|R,cColor:3:0.0.0|R,W11|"
    XLS = XLS & "S,TC|S,cLabel:End=S:3:Trans|S,cColor:3:0.0.0|S,W11|"
    XLS = XLS & "T,$Int|T,cLabel:End=T:3:LO Total Unit Cost|T,cColor:3:112.173.71|T,cFormula:2:=U2/$K2:$IntB|T,TC|T,W14|"
    XLS = XLS & "U,SIntE|U,cLabel:End=U:3:Group LO Total Unit Cost|U,cColor:3:204.102.0|U,TC|U,HH|"
    XLS = XLS & "V,$Int|V,cLabel:End=V:3:LO RV|V,cColor:3:112.173.71|V,cFormula:2:=W2/$K2:$IntB|V,TC|V,W11|"
    XLS = XLS & "W,SIntE|W,cLabel:End=W:3:Group LO RV|W,cColor:3:204.102.0|W,TC|W,HH|"
    XLS = XLS & "X,P1|X,cLabel:End=X:3:LO RV  %|X,cColor:3:112.173.71|X,FColor:0.0.0|X,cFormula:2:=V2/T2:P1B|X,TC|X,W07|"
    XLS = XLS & "Y,$Int|Y,cLabel:End=Y:3:LO Payment|Y,cColor:3:112.173.71|Y,cFormula:2:=Z2/$K2:$IntB|Y,TC|Y,W07|"
    XLS = XLS & "Z,SIntE|Z,cLabel:End=Z:3:Group LO Payment|Z,cColor:3:204.102.0|Z,TC|Z,HH|"
    XLS = XLS & "AA,P3|AA,cLabel:End=AA:3:LO LRF|AA,cColor:3:112.173.71|AA,FColor:0.0.0|AA,cFormula:2:=Y2/T2:P3B|AA,TC|AA,W06|"
    XLS = XLS & "AB,TC|AB,cLabel:End=AB:3:LO SWAP Type|AB,cColor:3:112.173.71|AB,W07|"
    XLS = XLS & "AC,P2|AC,cLabel:End=AC:3:LO SWAP Rate|AC,cColor:3:112.173.71|AC,FColor:0.0.0|AC,cFormula:2:=AD2/$K2:P1B|AC,TC|AC,W07|"
    XLS = XLS & "AD,SE|AD,cLabel:End=AD:3:Group LO Swap|AD,cColor:3:204.102.0|AD,FColor:0.0.0|AD,P2|AD,HH|"
    XLS = XLS & "AE,P2|AE,cLabel:End=AE:3:LO Spread|AE,cColor:3:112.173.71|AE,FColor:0.0.0|AE,cFormula:2:=AF2/$K2:P2B|AE,TC|AE,W07|"
    XLS = XLS & "AF,SE|AF,P2|AF,cLabel:End=AF:3:Group LO Spread|AF,cColor:3:204.102.0|AF,FColor:0.0.0|AF,TC|AF,HH|"
    XLS = XLS & "AG,P2|AG,cLabel:End=AG:3:LO Yield|AG,cColor:3:112.173.71|AG,FColor:0.0.0|AG,cFormula:2:=AH2/$K2:P2B|AG,TC|AG,W07|"
    XLS = XLS & "AH,SE|AH,P2|AH,cLabel:End=AH:3:Group LO Yield|AH,cColor:3:204.102.0|AH,FColor:0.0.0|AH,HH|"
    XLS = XLS & "AI,%R|AI,cLabel:End=AI:3:LO Bonus Depr|AI,cColor:3:112.173.71|AI,cFormula:2:=AJ2/$K2:%B|AI,TC|AI,W06|"
    XLS = XLS & "AJ,SE|AJ,%R|AJ,cLabel:End=AJ:3:Group LO Bonus|AJ,cColor:3:204.102.0|AJ,HH|"
    XLS = XLS & "AK,$Int|AK,cLabel:End=AK:3:LO Fee|AK,cColor:3:112.173.71|AK,cFormula:2:=AL2/$K2:$IntB|AK,TC|AK,W07|"
    XLS = XLS & "AL,SIntE|AL,cLabel:End=AL:3:Group LO Fee|AL,cColor:3:204.102.0|AL,HH|"
    XLS = XLS & "AM,P1|AM,cLabel:End=AM:3:LO Fee %|AM,cColor:3:112.173.71|AM,FColor:0.0.0|AM,cFormula:2:=AK2/K2:P1B|AM,TC|AM,W09|"
    XLS = XLS & "AN,$Int|AN,cLabel:End=AN:3:LO Term Penalty|AN,cColor:3:112.173.71|AN,cFormula:2:=AO2/$K2:$IntB|AN,TC|AN,W09|"
    XLS = XLS & "AO,SIntE|AO,cLabel:End=AO:3:Group LO Term Penalty|AO,cColor:3:204.102.0|AO,TC|AO,HH|"
    XLS = XLS & "AP,P1|AP,cLabel:End=AP:3:LO Term Penalty %|AP,cColor:3:112.173.71|AP,FColor:0.0.0|AP,cFormula:2:=AN2/K2:P1B|AP,TC|AP,W09|"
    XLS = XLS & "AQ,$Int|AQ,cLabel:End=AQ:3:LO Fee Net Term Penalty|AQ,cColor:3:112.173.71|AQ,cFormula:2:=AR2/$K2:$IntB|AQ,TC|AQ,W11|"
    XLS = XLS & "AR,SIntE|AR,cLabel:End=AR:3:Group LO Fee Net Term Penalty|AR,cColor:3:204.102.0|AR,TC|AR,HH|"
    XLS = XLS & "AS,P1|AS,cLabel:End=AS:3:LO Fee Net Term Penalty %|AS,cColor:3:112.173.71|AS,FColor:0.0.0|AS,cFormula:2:=AQ2/K2:P1B|AS,TC|AS,W08|"
    XLS = XLS & "AT,$Int|AT,cLabel:End=AT:3:LO Net Per Diem|AT,cColor:3:112.173.71|AT,cFormula:2:=AU2/$K2:$IntB|AT,W11|"
    XLS = XLS & "AU,SIntE|AU,cLabel:End=AU:3:Group LO Net Per Diem|AU,cColor:3:204.102.0|AU,HH|"
    XLS = XLS & "AV,P1|AV,cLabel:End=AV:3:LO Net Per Diem %|AV,cColor:3:112.173.71|AV,FColor:0.0.0|AV,cFormula:2:=AT2/K2:P1B|AV,TC|AV,W07|"
    XLS = XLS & "AW,$Int|AW,cLabel:End=AW:3:LO UTA (Rebate)|AW,cColor:3:112.173.71|AW,cFormula:2:=AX2/$K2:$IntB|AW,TC|AW,W11|"
    XLS = XLS & "AX,SIntE|AX,cLabel:End=AX:3:Group LO UTA (Rebate)|AX,cColor:3:204.102.0|AX,TC|AX,HH|"
    XLS = XLS & "AY,P1|AY,cLabel:End=AY:3:LO UTA (Rebate) %|AY,cColor:3:112.173.71|AY,FColor:0.0.0|AY,cFormula:2:=AW2/K2:P1B|AY,TC|AY,W07|"
    XLS = XLS & "AZ,$Int|AZ,cLabel:End=AZ:3:LO Tota lFee|AZ,cColor:3:112.173.71|AZ,cFormula:2:=BA2/$K2:$IntB|AZ,TC|AZ,W11|"
    XLS = XLS & "BA,SIntE|BA,cLabel:End=BA:3:Group LO Total Fee|BA,cColor:3:204.102.0|BA,TC|BA,HH|"
    XLS = XLS & "BB,P1|BB,cLabel:End=BB:3:LO Total Fee %|BB,cColor:3:112.173.71|BB,FColor:0.0.0|BB,cFormula:2:=AZ2/K2:P1B|BB,TC|BB,W11|"
    XLS = XLS & "BC,$Int|BC,cLabel:End=BC:3:Synd Tractor Only Cost|BC,cFormula:2:=BD2/$K2:$IntB|BC,TC|"
    XLS = XLS & "BD,SIntE|BD,cLabel:End=BD:3:Group Synd Tractor Only Cost|BD,cColor:3:204.102.0|BD,TC|BD,HH|"
    XLS = XLS & "BE,$Int|BE,cLabel:End=BE:3:Synd Child Parts|BE,cFormula:2:=BF2/$K2:$IntB|BE,TC|BE,W11|"
    XLS = XLS & "BF,SIntE|BF,cLabel:End=BF:3:Group Synd Child Parts|BF,cColor:3:204.102.0|BF,TC|BF,HH|"
    XLS = XLS & "BG,$Int|BG,cLabel:End=BG:3:Synd Total Vendor Cost|BG,cFormula:2:=BH2/$K2:$IntB|BG,TC|BG,W11|"
    XLS = XLS & "BH,SIntE|BH,cLabel:End=BH:3:Group Synd Total Vendor Cost|BH,cColor:3:204.102.0|BH,TC|BH,HH|"
    XLS = XLS & "BI,$Int|BI,cLabel:End=BI:3:Synd RV|BI,cFormula:2:=BJ2/$K2:$IntB|BI,TC|BI,W12|"
    XLS = XLS & "BJ,SIntE|BJ,cLabel:End=BJ:3:Group Synd RV|BJ,cColor:3:204.102.0|BJ,TC|BJ,HH|"
    XLS = XLS & "BK,P1|BK,cLabel:End=BK:3:Synd RV Tractor Only %|BK,FColor:0.0.0|BK,cFormula:2:=BI2/BC2:P1B|BK,TC|BK,W12|"
    XLS = XLS & "BL,P1|BL,cLabel:End=BL:3:Synd RV Total OEC %|BL,FColor:0.0.0|BL,cFormula:2:=BI2/BG2:P1B|BL,TC|BL,W11|"
    XLS = XLS & "BM,$Int|BM,cLabel:End=BM:3:Payment|BM,cFormula:2:=BN2/$K2:$IntB|BM,W11|"
    XLS = XLS & "BN,SIntE|BN,cLabel:End=BN:3:Group Payment|BN,cColor:3:204.102.0|BN,HH|"
    XLS = XLS & "BO,P3|BO,cLabel:End=BO:3:LRF|BO,FColor:0.0.0|BO,cFormula:2:=BM2/BG2:P3B|BO,TC|BO,W11|"
    XLS = XLS & "BP,TC|BP,cLabel:End=BP:3:Synd SWAP Type|BP,W11|"
    XLS = XLS & "BQ,P2|BQ,cLabel:End=BQ:3:Synd SWAP Rate|BQ,cFormula:2:=BR2/$K2:P2B|BQ,TC|BQ,W11|"
    XLS = XLS & "BR,SE|BR,P2|BR,cLabel:End=BR:3:Group Synd SWAP Rate|BR,cColor:3:204.102.0|BR,TC|BR,HH|"
    XLS = XLS & "BS,P2|BS,cLabel:End=BS:3:Synd Spread|BS,cFormula:2:=BT2/$K2:P2B|BS,TC|BS,W06|"
    XLS = XLS & "BT,SE|BT,P2|BT,cLabel:End=BT:3:Group Synd Spread|BT,cColor:3:204.102.0|BT,TC|BT,HH|"
    XLS = XLS & "BU,P2|BU,cLabel:End=BU:3:Synd Yield |BU,cFormula:2:=BV2/$K2:P2B|BU,TC|BU,W06|"
    XLS = XLS & "BV,SE|BV,P2|BV,cLabel:End=BV:3:Group Synd Yield |BV,cColor:3:204.102.0|BV,TC|BV,HH|"
    XLS = XLS & "BW,%R|BW,cLabel:End=BW:3:Synd Bonus Depr|BW,cFormula:2:=BX2/$K2:%B|BW,TC|BW,W06|"
    XLS = XLS & "BX,SE|BX,%R|BX,cLabel:End=BX:3:Group Synd Bonus Depr|BX,cColor:3:204.102.0|BX,TC|BX,HH|"
    XLS = XLS & "BY,$Int|BY,cLabel:End=BY:3:Synd Fee|BY,cFormula:2:=BZ2/$K2:$IntB|BY,TC|BY,W08|"
    XLS = XLS & "BZ,SIntE|BZ,cLabel:End=BZ:3:Group Synd Fee|BZ,cColor:3:204.102.0|BZ,TC|BZ,HH|"
    XLS = XLS & "CA,P1|CA,cLabel:End=CA:3:Synd Fee %|CA,FColor:0.0.0|CA,cFormula:2:=BY2/BG2:P1B|CA,TC|CA,W11|"
    XLS = XLS & "CB,$Int|CB,cLabel:End=CB:3:Synd Term Penalty|CB,cFormula:2:=CC2/$K2:$IntB|CB,TC|CB,W08|"
    XLS = XLS & "CC,SIntE|CC,cLabel:End=CC:3:Group Synd Term Penalty|CC,cColor:3:204.102.0|CC,TC|CC,HH|"
    XLS = XLS & "CD,P1|CD,cLabel:End=CD:3:Synd Term Penalty %|CD,FColor:0.0.0|CD,cFormula:2:=CB2/BG2:P1B|CD,W11|"
    XLS = XLS & "CE,$Int|CE,cLabel:End=CE:3:Synd Fee Net Term Penalty|CE,cFormula:2:=CF2/$K2:$IntB|CE,TC|CE,W09|"
    XLS = XLS & "CF,SIntE|CF,cLabel:End=CF:3:Group Synd Fee Net Term Penalty|CF,cColor:3:204.102.0|CF,TC|CF,HH|"
    XLS = XLS & "CG,P1|CG,cLabel:End=CG:3:Synd Fee Net Term Penalty %|CG,FColor:0.0.0|CG,cFormula:2:=CF2/BG2:P1B|"
    XLS = XLS & "CH,$Int|CH,cLabel:End=CH:3:Synd Net Per Diem|CH,cFormula:2:=CI2/$K2:$IntB|CH,TC|CH,W09|"
    XLS = XLS & "CI,SIntE|CI,cLabel:End=CI:3:Group Synd Net Per Diem|CI,cColor:3:204.102.0|CI,TC|CI,HH|"
    XLS = XLS & "CJ,P1|CJ,cLabel:End=CJ:3:Synd Net Per Diem %|CJ,FColor:0.0.0|CJ,cFormula:2:=CH2/BG2:P1B|CJ,W07|"
    XLS = XLS & "CK,$Int|CK,cLabel:End=CK:3:Synd UTA (Rebate)|CK,cFormula:2:=CL2/$K2:$IntB|CK,TC|CK,W07|"
    XLS = XLS & "CL,SIntE|CL,cLabel:End=CL:3:Group Synd UTA (Rebate)|CL,cColor:3:204.102.0|CL,TC|CL,HH|"
    XLS = XLS & "CM,P1|CM,cLabel:End=CM:3:Synd UTA (Rebate) %|CM,FColor:0.0.0|CM,cFormula:2:=CK2/BG2:P1B|CM,W11|"
    XLS = XLS & "CN,$Int|CN,cLabel:End=CN:3:Synd Total Fee (w/Strips)|CN,cFormula:2:=CO2/$K2:$IntB|CN,TC|CN,W07|"
    XLS = XLS & "CO,SIntE|CO,cLabel:End=CO:3:Group Synd Total Fee|CO,cColor:3:204.102.0|CO,TC|CO,HH|"
    XLS = XLS & "CP,P1|CP,cLabel:End=CP:3:Synd Total Fee (w/Strips) % |CP,FColor:0.0.0|CP,cFormula:2:=CN2/BG2:P1B|CP,TC|CP,W08|"
    XLS = XLS & "A,WRAP|"
    TitleX = ""
    TitleDateX = ""
    If IsDate(StartDateRange) Then TitleDateX = TitleDateX & StartDateRange & " "
    If IsDate(StartDateRange) And IsDate(EndDateRange) Then TitleDateX = TitleDateX & "~ "
    If IsDate(EndDateRange) Then TitleDateX = TitleDateX & EndDateRange & " "
    If Nz(ClientGroupIDi, 0) <> 0 Then TitleX = TitleX & "ClientGroup: " & DLookup("ClientGroupName", "ClientGroups", "ClientGroupID=" & ClientGroupIDi & "") & ""
    If Nz(BankIDi, 0) <> 0 Then TitleX = TitleX & "Bank " & DLookup("BankShortName", "Banks", "BankId=" & BankIDi) & " "

    ' If IsDate(StartDateRange) Then SqlDate = "AND FundDate >=#" & StartDateRange & "# "
    ' If IsDate(EndDateRange) Then SqlDate = SqlDate & "AND FundDate <=#" & EndDateRange & "# "
    If Nz(ClientGroupIDi, 0) <> 0 Then SqlClient = "AND ClientGroupID=" & ClientGroupIDi & " "
    ' If Nz(BankIDi, 0) <> 0 Then SqlBank = "AND BankId=" & BankIDi & " "
    ' If IsDate(EndDateRange) Or IsDate(StartDateRange) Then SqlNull = "Or FundDate Is Null "
    SqlDate = SqlDate & SqlClient & SqlBank
    SqlNull = SqlNull & SqlClient & SqlBank
    If FundStatus = 2 Then SQLz = SQLz & " WHERE NOT Assigned_Indicator is Null " & SqlDate Else SQLz = SQLz & " WHERE Assigned_Indicator = 'F' " & SqlDate
    Call MessageUser("Update", "TMT-TRF For All Groups", "Setting Funded, just a sec...")
    Call SendReport2Excel(SQLz, "TMT-TFR", "TMT-TFR " & TitleX & TitleDateX, True, XLS, 0, 0, 0, , , "C:\Faas\" & FaUserName & "\Temp\TMT-TFR-Report.xlsx")
    Call MessageUser("Close")
End Function
Public Function TFRExcel(Optional StartDateRange As String, Optional EndDateRange As String, Optional ClientGroupIDi As Integer, Optional BankIDi As Integer)
    If FAUserAccessGroup Like "*ACCT*" Or FAUserAccessGroup Like "*ITT*" Or FAUserAccessGroup Like "*TMT*" Then
        'Call MakeTFRWorkTable
    Else
        MsgBox "Sorry you do NOT have access to this report.", vbCritical, AtlaasVersion
        Exit Function
    End If
    Dim SQLz As String
    Dim SQLy As String
    Dim SQLx As String
    Dim SqlA As String
    Dim Sqlb As String
    Dim SqlDate, SqlClient, SqlBank, SqlNull, DateRangeX As String
    Dim Go1, Go2, Go3, Go4, NumRpts, GrpRpt As Integer
    Dim IsLastRpt As Boolean
    '   Dim RST    As DAO.Recordset
    Dim XLS, TitleX, TitleDateX As String
    Go1 = 0
    Go2 = 0
    Go3 = 0
    Go4 = 0
    Call MessageUser("Open", "Setting up TRF", "This will take a few moments")

    SQLx = "Select ClientGroup" '- Col:A
    SQLx = SQLx & " ,AccountName" '- Col:B
    SQLx = SQLx & " ,Schedule" '- Col:C
    SQLx = SQLx & " ,UnitGroup" '- Col:D
    SQLx = SQLx & " ,FAID" '- Col:E
    SQLx = SQLx & " ,LPAFYear" '- Col:F
    SQLx = SQLx & " ,Assigned_Indicator" '- Col:G
    SQLx = SQLx & " ,FundDate" '- Col:H
    SQLx = SQLx & " ,Stillactive" '- Col:I
    SQLx = SQLx & " ,SchIn" '- Col:J
    SQLx = SQLx & " ,MaxEstDelDate" '- Col:K
    SQLx = SQLx & " ,BLCD" '- Col:L
    SQLx = SQLx & " ,LeaseType" '- Col:M
    SQLx = SQLx & " ,SLB" '- Col:N
    SQLx = SQLx & " ,LeaseExp" '- Col:O
    SQLx = SQLx & " ,Bank" '- Col:P
    SQLx = SQLx & " ,LeaseTermPrimary" '- Col:Q
    SQLx = SQLx & " ,MilesAllowedTerm" '- Col:R
    SQLx = SQLx & " ,MilesAllowed" '- Col:S
    SQLx = SQLx & " ,ODORange" '- Col:T
    SQLx = SQLx & " ,UnitCount" '- Col:U
    SQLx = SQLx & " ,DescAbbrv" '- Col:V
    SQLx = SQLx & " ,Myear" '- Col:W
    SQLx = SQLx & " ,Make" '- Col:X
    SQLx = SQLx & " ,Model" '- Col:Y
    SQLx = SQLx & " ,Type" '- Col:Z
    SQLx = SQLx & " ,SubType" '- Col:AA
    SQLx = SQLx & " ,Axle" '- Col:AB
    SQLx = SQLx & " ,OECPerUnit" '- Col:AC
    SQLx = SQLx & " ,GroupOEC" '- Col:AD
    SQLx = SQLx & " ,VendorCostUnit" '- Col:AE
    SQLx = SQLx & " ,TotalVendorCost" '- Col:AF
    SQLx = SQLx & " ,SyndFee" '- Col:AG
    SQLx = SQLx & " ,SyndFeePct" '- Col:AH
    SQLx = SQLx & " ,TermPenalty" '- Col:AI
    SQLx = SQLx & " ,TermPenaltyPct" '- Col:AJ
    SQLx = SQLx & " ,InterimExp" '- Col:AK
    SQLx = SQLx & " ,NetSyndFee" '- Col:AL
    SQLx = SQLx & " ,NetSyndPct" '- Col:AM
    SQLx = SQLx & " ,StripRent" '- Col:AN
    SQLx = SQLx & " ,StripRentPct" '- Col:AO
    SQLx = SQLx & " ,PDRent" '- Col:AP
    SQLx = SQLx & " ,PDRentPct" '- Col:AQ
    SQLx = SQLx & " ,PDDaysAvg" '- Col:AR
    SQLx = SQLx & " ,UTA" '- Col:AS
    SQLx = SQLx & " ,UTAPct" '- Col:AT
    SQLx = SQLx & " ,NPV" '- Col:AU
    SQLx = SQLx & " ,NPVPct" '- Col:AV
    SQLx = SQLx & " ,TotalFee" '- Col:AW
    SQLx = SQLx & " ,TotalFeePct" '- Col:AX
    SQLx = SQLx & " ,Payment" '- Col:AY
    SQLx = SQLx & " ,PmtAmt" '- Col:AZ
    SQLx = SQLx & " ,LRF" '- Col:BA
    SQLx = SQLx & " ,FMSS" '- Col:BB
    SQLx = SQLx & " ,PmtAmtFMSS" '- Col:BC
    SQLx = SQLx & " ,RVPerUnit" '- Col:BD
    SQLx = SQLx & " ,RVPct" '- Col:BE
    SQLx = SQLx & " ,SWAPType" '- Col:BF
    SQLx = SQLx & " ,Yield" '- Col:BG
    SQLx = SQLx & " ,IndexRate" '- Col:BH
    SQLx = SQLx & " ,Spread" '- Col:BI
    SQLx = SQLx & " ,SpreadIRR" '- Col:BJ
    SQLx = SQLx & " ,YieldIRR" '- Col:BK
    SQLx = SQLx & " ,Bonus" '- Col:BL
    SQLx = SQLx & " ,DealCommission" '- Col:BM
    SQLx = SQLx & " ,SyndFeeProj" '- Col:BN
    SQLx = SQLx & " ,LoSyndPct" '- Col:BO
    SQLx = SQLx & " ,TermPenaltyProj" '- Col:BP
    SQLx = SQLx & " ,LoTermPenaltyPct" '- Col:BQ
    SQLx = SQLx & " ,LoPDUnit" '- Col:BR
    SQLx = SQLx & " ,LoPDPct" '- Col:BS
    SQLx = SQLx & " ,LoUATUnit" '- Col:BT
    SQLx = SQLx & " ,LoUTAPct" '- Col:BU
    SQLx = SQLx & " ,SpreadProj" '- Col:BV
    SQLx = SQLx & " ,NPVProj" '- Col:BW
    SQLx = SQLx & " ,BankContract" '- Col:BX
    SQLx = SQLx & " ,RVLOUnit" '- Col:BY
    SQLx = SQLx & " ,RVLOUnitPct" '- Col:BZ
    SQLx = SQLx & " From TFR_2 "

    XLS = XLS & "|"
    XLS = XLS & "A,DD|A,cLabel:End=A:3:Client Group|A,TR|"
    XLS = XLS & "B,cLabel:End=B:3:Client|B,BL|"
    XLS = XLS & "C,cLabel:End=C:3:Sch|C,TL|C,W07|"
    XLS = XLS & "D,cLabel:End=D:3:Grp|D,TC|D,W06|"
    XLS = XLS & "E,cLabel:End=E:3:FAID|E,TC|E,W09|"
    XLS = XLS & "F,cLabel:End=F:3:LPAF FYR|F,cColor:3:0.0.0|F,TC|F,W06|"
    XLS = XLS & "G,cLabel:End=G:3:Fund Status|G,cColor:3:0.0.0|G,BC|G,W06|"
    XLS = XLS & "H,cLabel:End=H:3:Fund Date|H,cColor:3:0.0.0|H,TC|H,W10|"
    XLS = XLS & "I,cLabel:End=I:3:Still Active Units|I,cColor:3:84.130.53|I,BC|I,W07|"
    XLS = XLS & "J,cLabel:End=J:3:Sch IN (Signed)|J,TC|J,W10|"
    XLS = XLS & "K,cLabel:End=K:3:Est Delv (max)|K,TC|K,W10|"
    XLS = XLS & "L,cLabel:End=L:3:BLCD|L,BC|L,W10|"
    XLS = XLS & "M,cLabel:End=M:3:Lease Type|M,cColor:3:48.84.150|M,TC|M,W10|"
    XLS = XLS & "N,cLabel:End=N:3:SLP|N,cColor:3:48.84.150|N,BC|N,W05|"
    XLS = XLS & "O,cLabel:End=O:3:Lease Expiry|O,TC|O,W10|"
    XLS = XLS & "P,cLabel:End=P:3:Bank|P,BL|"
    XLS = XLS & "Q,cLabel:End=Q:3:Term Mths|Q,BC|Q,W06|"
    XLS = XLS & "R,Sint|R,cLabel:End=R:3:Miles Term|R,cColor:3:0.0.40|R,TC|"
    XLS = XLS & "S,Sint|S,cLabel:End=S:3:Miles Pyear|S,cColor:3:0.0.50|S,TC|"
    XLS = XLS & "T,cLabel:End=T:3:ODO Range (Avg)|T,cColor:3:0.0.50|T,TC|T,W14|"
    XLS = XLS & "U,cLabel:End=U:3:Unit Count|U,BC|U,W07|"
    XLS = XLS & "V,cLabel:End=V:3:Unit Desc|V,W25|"
    XLS = XLS & "W,cLabel:End=W:3:Model Year|W,W07|"
    XLS = XLS & "X,cLabel:End=X:3:Unit Make|X,W15|"
    XLS = XLS & "Y,cLabel:End=Y:3:Unit Model|Y,W15|"
    XLS = XLS & "Z,cLabel:End=Z:3:Unit Type|"
    XLS = XLS & "AA,cLabel:End=AA:3:Sub Type|"
    XLS = XLS & "AB,cLabel:End=AB:3:Axle|AB,TC|"
    XLS = XLS & "AC,$Int|AC,cLabel:End=AC:3:OEC Unit (avg)|"
    XLS = XLS & "AD,$Int|AD,cLabel:End=AD:3:Group OEC|"
    XLS = XLS & "AE,$Int|AE,cLabel:End=AE:3:Vendor Cost (avg)|AE,cColor:3:0.0.50|"
    XLS = XLS & "AF,$Int|AF,cLabel:End=AF:3:Group Vendor Cost|AF,cColor:3:0.0.50|"
    XLS = XLS & "AG,$Int|AG,cLabel:End=AG:3:Gross Synd Fee|"
    XLS = XLS & "AH,P2|AH,cLabel:End=AH:3:Gross Synd %|AH,FColor:47.117.181|AH,W07|"
    XLS = XLS & "AI,$Int|AI,cLabel:End=AI:3:Term Penalty|"
    XLS = XLS & "AJ,P2|AJ,cLabel:End=AJ:3:Term Penalty %|AJ,FColor:47.117.181|AJ,W07|"
    XLS = XLS & "AK,$IntE|AK,cLabel:End=AK:3:Interim Exp|"
    XLS = XLS & "AL,$IntE|AL,cLabel:End=AL:3:Net Synd Fee|AL,W09|"
    XLS = XLS & "AM,P2|AM,cLabel:End=AM:3:Net Synd %|AM,FColor:47.117.181|AM,W07|"
    XLS = XLS & "AN,$IntE|AN,cLabel:End=AN:3:Strip Rent|"
    XLS = XLS & "AO,P2|AO,cLabel:End=AO:3:Strip Rent %|AO,FColor:47.117.181|"
    XLS = XLS & "AP,$IntE|AP,cLabel:End=AP:3:PerDiem Group|"
    XLS = XLS & "AQ,P2|AQ,cLabel:End=AQ:3:Per Diem %|AQ,FColor:47.117.181|"
    XLS = XLS & "AR,cLabel:End=AR:3:PerDiem Days (avg)|AR,BC|"
    XLS = XLS & "AS,$IntE|AS,cLabel:End=AS:3:UTA|"
    XLS = XLS & "AT,P2|AT,cLabel:End=AT:3:UTA %|AT,FColor:47.117.181|"
    XLS = XLS & "AU,$IntE|AU,cLabel:End=AU:3:FMSS NPV|"
    XLS = XLS & "AV,P2|AV,cLabel:End=AV:3:NPV %|AV,FColor:47.117.181|"
    XLS = XLS & "AW,$IntE|AW,cLabel:End=AW:3:Total Fee|"
    XLS = XLS & "AX,P2|AX,cLabel:End=AX:3:Total Fee %|AX,FColor:47.117.181|"
    XLS = XLS & "AY,$IntE|AY,cLabel:End=AY:3:Payment (sum)|AY,cColor:3:51.63.79|"
    XLS = XLS & "AZ,$R|AZ,cLabel:End=AZ:3:Payment (avg)|"
    XLS = XLS & "BA,P2|BA,cLabel:End=BA:3:LRF|"
    XLS = XLS & "BB,$IntE|BB,cLabel:End=BB:3:FMSS|"
    XLS = XLS & "BC,$IntE|BC,cLabel:End=BC:3:PMT FMSS|"
    XLS = XLS & "BD,$IntE|BD,cLabel:End=BD:3:RV Unit (avg)|"
    XLS = XLS & "BE,P2|BE,cLabel:End=BE:3:RV Unit %|BE,FColor:84.130.53|BE,TC|"
    XLS = XLS & "BF,cLabel:End=BF:3:SWAP Type|BF,cColor:3:0.0.0|BF,BC|"
    XLS = XLS & "BG,P2|BG,cLabel:End=BG:3:Yield MISF|BG,cColor:3:0.0.0|BG,TC|"
    XLS = XLS & "BH,P2|BH,cLabel:End=BH:3:Index Rate|BH,cColor:3:55.86.35|BH,TC|"
    XLS = XLS & "BI,P2|BI,cLabel:End=BI:3:Spread|BI,cColor:3:55.86.35|BI,TC|"
    XLS = XLS & "BJ,P2|BJ,cLabel:End=BJ:3:Spread IRR-PT|BJ,cColor:3:84.130.53|BJ,TC|"
    XLS = XLS & "BK,P2|BK,cLabel:End=BK:3:Yield IRR-PT|BK,cColor:3:84.130.53|BK,TC|"
    XLS = XLS & "BL,P2|BL,cLabel:End=BL:3:Bonus Depr|BL,TC|"
    XLS = XLS & "BM,P2|BM,cLabel:End=BM:3:Deal Comm|BM,TC|"
    XLS = XLS & "BN,$IntE|BN,cLabel:End=BN:3:LO Synd Fee|BN,cColor:3:150.150.150|"
    XLS = XLS & "BO,P2|BO,cLabel:End=BO:3:LO SyndFee %|BO,cColor:3:170.170.170|BO,FColor:47.117.181|BO,TC|"
    XLS = XLS & "BP,$IntE|BP,cLabel:End=BP:3:LO Term Penalty|BP,cColor:3:170.170.170|"
    XLS = XLS & "BQ,P2|BQ,cLabel:End=BQ:3:LO Term %|BQ,cColor:3:150.150.150|BQ,FColor:47.117.181|"
    XLS = XLS & "BR,$IntE|BR,cLabel:End=BR:3:LO PD Unit|BR,cColor:3:170.170.170|"
    XLS = XLS & "BS,P2|BS,cLabel:End=BS:3:LO PD %|BS,cColor:3:150.150.150|BS,FColor:47.117.181|"
    XLS = XLS & "BT,$IntE|BT,cLabel:End=BT:3:LO UTA Unit|BT,cColor:3:170.170.170|"
    XLS = XLS & "BU,P2|BU,cLabel:End=BU:3:LO UTA %|BU,cColor:3:150.150.150|BU,FColor:47.117.181|"
    XLS = XLS & "BV,P2|BV,cLabel:End=BV:3:LO Spread|BV,cColor:3:170.170.170|"
    XLS = XLS & "BW,$IntE|BW,cLabel:End=BW:3:LO NPV|BW,cColor:3:150.150.150|"
    XLS = XLS & "BX,cLabel:End=BX:3:Bank Contract|"
    XLS = XLS & "BY,$Int|BY,cLabel:End=BY:3:LO RV Unit|"
    XLS = XLS & "BZ,P2|BZ,cLabel:End=BZ:3:LO  RV %|BZ,FColor:47.117.181|"
    XLS = XLS & "A,WRAP|"

    TitleX = ""
    TitleDateX = ""
    If IsDate(StartDateRange) Then TitleDateX = TitleDateX & StartDateRange & " "
    If IsDate(StartDateRange) And IsDate(EndDateRange) Then TitleDateX = TitleDateX & "~ "
    If IsDate(EndDateRange) Then TitleDateX = TitleDateX & EndDateRange & " "
    If ClientGroupIDi <> 0 Then TitleX = TitleX & "Client " & DLookup("ClientGroupName", "ClientGroups", "ClientGroupID=" & ClientGroupIDi)
    If Nz(BankIDi, 0) <> 0 Then TitleX = TitleX & "Bank " & DLookup("BankShortName", "Banks", "BankId=" & BankIDi) & " "

    If IsDate(StartDateRange) Then SqlDate = "AND FundDate >=#" & StartDateRange & "# "
    If IsDate(EndDateRange) Then SqlDate = SqlDate & "AND FundDate <=#" & EndDateRange & "# "
    If Nz(ClientGroupIDi, 0) <> 0 Then SqlClient = "AND ClientGroupID=" & ClientGroupIDi & " "
    If Nz(BankIDi, 0) <> 0 Then SqlBank = "AND BankId=" & BankIDi & " "
    If IsDate(EndDateRange) Or IsDate(StartDateRange) Then SqlNull = "Or FundDate Is Null "
    SqlDate = SqlDate & SqlClient & SqlBank
    SqlNull = SqlNull & SqlClient & SqlBank

Forecast:
    SQLy = SQLx & " WHERE Assigned_Indicator = 'P' " & SqlClient & SqlBank
    SQLy = SQLy & " ORDER BY AccountName, CDbl(NumInString(Left(Schedule,3))) desc, Schedule, UnitGroup"
    'If DCount("sGrpID", "TFR_Work", "Assigned_Indicator='P' " & SqlDate) >= 1 Then Go4 = 1
Inventory:
    Sqlb = SQLx & " WHERE Assigned_Indicator= 'I' " & SqlClient & SqlBank
    Sqlb = Sqlb & " ORDER BY AccountName, CDbl(NumInString(Left(Schedule,3))) desc, Schedule, UnitGroup"
    'If DCount("sGrpID", "TFR_Work", "Assigned_Indicator='I' " & SqlDate) >= 1 Then Go3 = 1
Hold:
    SqlA = SQLx & " WHERE Assigned_Indicator= 'H' " & SqlDate
    SqlA = SqlA & " ORDER BY AccountName, CDbl(NumInString(Left(Schedule,3))) desc, Schedule, UnitGroup"
    'If DCount("sGrpID", "TFR_Work", "Assigned_Indicator='H' " & SqlDate) >= 1 Then Go2 = 1
Final:
    'SQLx = SQLx & " WHERE Assigned_Indicator = 'F' " & SqlDate
    'If DCount("sGrpID", "TFR_Work", "Assigned_Indicator= 'F' " & SqlDate) >= 1 Then Go1 = 1
    SQLx = SQLx & " WHERE Assigned_Indicator <> '' " & SqlDate
    SQLx = SQLx & " ORDER BY AccountName, CDbl(NumInString(Left(Schedule,3))) desc, Schedule, UnitGroup"
    If DCount("sGrpID", "TFR_Work", "Assigned_Indicator<> '' " & SqlDate) >= 1 Then Go1 = 1

    NumRpts = (Go1 + Go2 + Go3 + Go4)
    If NumRpts = 0 Then
        MsgBox "Sorry, but your request returns NO data", vbCritical, AtlaasVersion
        Call MessageUser("Close")
        Exit Function
    End If
    If NumRpts <= 1 Then IsLastRpt = True Else IsLastRpt = False
    If Go4 = 1 Then
        Call MessageUser("Update", "TRF For Forecast Groups", "Setting Forecast Tab, Just a bit...")
        Call SendReport2Excel(SQLy, "Forecast", "Forecast " & TitleX, True, XLS, 0, IsLastRpt, Go4, IsLastRpt, , "C:\Faas\" & FaUserName & "\Temp\TFRReport.xlsx")
        If NumRpts > 1 Then NumRpts = NumRpts - 1
    End If

    If NumRpts <= 1 Then IsLastRpt = True Else IsLastRpt = False
    If Go3 = 1 Then
        Call MessageUser("Update", "TRF For Inventory Groups", "Setting Inventory Tab, working...")
        Call SendReport2Excel(Sqlb, "Inventory", "Inventory " & TitleX, True, XLS, 0, IsLastRpt, Go4 + Go3, IsLastRpt, , "C:\Faas\" & FaUserName & "\Temp\TFRReport.xlsx")
        If NumRpts > 1 Then NumRpts = NumRpts - 1
    End If
    If NumRpts <= 1 Then IsLastRpt = True Else IsLastRpt = False
    If Go2 = 1 Then
        Call MessageUser("Update", "TRF For Hold Groups", "Making Inventory Tab, closer...")
        Call SendReport2Excel(SqlA, "FA Hold", "Hold " & TitleX, True, XLS, 0, IsLastRpt, Go4 + Go3 + Go2, IsLastRpt, , "C:\Faas\" & FaUserName & "\Temp\TFRReport.xlsx")
        If NumRpts > 1 Then NumRpts = NumRpts - 1
    End If
    If NumRpts <= 1 Then IsLastRpt = True Else IsLastRpt = False
    If Go1 = 1 Then
        Call MessageUser("Update", "TRF For All Groups", "Setting Funded, just a sec...")
        Call SendReport2Excel(SQLx, "TFR All", "All Statuses " & TitleX & TitleDateX, True, XLS, 0, 0, Go4 + Go3 + Go2 + Go1, IsLastRpt, , "C:\Faas\" & FaUserName & "\Temp\TFRReport.xlsx")
        'Call MessageUser("Update", "TRF For Funded Groups", "Setting Funded, just a sec...")
        'Call SendReport2Excel(SQLx, "Final", "Funded " & TitleX & TitleDateX, True, XLS, 0, IsLastRpt, Go4 + Go3 + Go2 + Go1, IsLastRpt, , "C:\Faas\" & FAUserName & "\Temp\TFRReport.xlsx")
    End If

    Call MessageUser("Close")
End Function
