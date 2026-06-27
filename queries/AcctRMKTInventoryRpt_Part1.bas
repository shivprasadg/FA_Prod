dbMemo "SQL" ="SELECT DISTINCT vw_Sch_Grp_Units_Remarketing.ClientShortName AS Client, vw_Sch_G"
    "rp_Units_Remarketing.AssetID, vw_Sch_Grp_Units_Remarketing.UnitNum, vw_Sch_Grp_U"
    "nits_Remarketing.VIN, LookUp_UnitDesc.DescriptionMed AS Description, Locations.L"
    "ocationName AS Location, [LAddrLine1] & \" \" & [LCity] & \" \" & [LState] & \" "
    "\" & [LZip] AS [Location Address], CCur(Nz([Rmkt].[BuyOutCost],0)) AS [BuyOut Co"
    "st], CCur(IIf([vw_Sch_Grp_Units_Remarketing].[ClientGroupId]=432 And [rmkt].[Buy"
    "OutDate]>#6/15/2017# And [rmkt].[BuyOutDate]<#7/5/2017#,Nz([TermPenalty],0),0)) "
    "AS Adjustment, CCur(Nz([Rmkt].[BuyOutCost],0))+CCur(IIf([vw_Sch_Grp_Units_Remark"
    "eting].[ClientGroupId]=432,Nz([TermPenalty],0),0)) AS [Final Cost], Rmkt.Invoice"
    "PaidDate, Rmkt.BuyOutDate, vw_Sch_Grp_Units_Remarketing.Status, Banks.BankShortN"
    "ame AS InterimBank, IIf([IFPaidDate]<=[forms]![acctmenu]![AsOfDate],0,CCur(Nz([R"
    "mkt.IFAmount],0))) AS InterimAmt, Rmkt.IFLoanNumber AS AccountNumber, Rmkt.IFDat"
    "e AS InterimDate, Rmkt.IFPaidDate AS InterimPaidDate, CCur(Nz([Final Cost],0)-Nz"
    "([Rmkt].[IFAmount],0)) AS FAAmount, IIf(IsNull([Rmkt].[IFAmount]),Null,Round(1-("
    "[Rmkt].[IFAmount]/[Final Cost]),3)) AS FAPercent, Clients.ClientShNm AS Buyer, R"
    "mkt.SoldPrice AS ActualSoldPrice, Units.unitoffleasedt AS ActualOffLease, Switch"
    "([Wholesale]=2,\"OEM Trade\",[Wholesale]=0,\"Retail\",[Wholesale]=1,\"Wholesale\""
    ") AS DSPN, vw_UnitInspectionLast.iPassedDate AS InspPassed\015\012FROM (((((vw_S"
    "ch_Grp_Units_Remarketing INNER JOIN (Units LEFT JOIN Locations ON Units.UnitGara"
    "geLocationRef = Locations.LocationID) ON vw_Sch_Grp_Units_Remarketing.AssetID = "
    "Units.UnitID) INNER JOIN Rmkt ON Units.UnitID = Rmkt.UnitRef) LEFT JOIN Banks ON"
    " Rmkt.IFBankID = Banks.BankId) LEFT JOIN Clients ON Rmkt.BuyerRef = Clients.clie"
    "ntID) LEFT JOIN LookUp_UnitDesc ON Rmkt.UnitRef = LookUp_UnitDesc.AssetID) LEFT "
    "JOIN vw_UnitInspectionLast ON Units.UnitID = vw_UnitInspectionLast.AssetID\015\012"
    "WHERE (((Rmkt.InvoicePaidDate) Is Null Or (Rmkt.InvoicePaidDate)>forms!acctmenu!"
    "AsOfDate) And ((Rmkt.BuyOutDate)<=forms!acctmenu!AsOfDate))\015\012ORDER BY vw_S"
    "ch_Grp_Units_Remarketing.ClientShortName, vw_Sch_Grp_Units_Remarketing.AssetID, "
    "vw_Sch_Grp_Units_Remarketing.UnitNum;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbInteger "RowHeight" ="345"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="-1"
Begin
    Begin
        dbText "Name" ="Location"
        dbInteger "ColumnWidth" ="2415"
        dbInteger "ColumnOrder" ="6"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Sch_Grp_Units_Remarketing.UnitNum"
        dbInteger "ColumnWidth" ="1620"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Sch_Grp_Units_Remarketing.VIN"
        dbInteger "ColumnWidth" ="2295"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Description"
        dbInteger "ColumnWidth" ="5295"
        dbInteger "ColumnOrder" ="5"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Sch_Grp_Units_Remarketing.AssetID"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.InvoicePaidDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="Client"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Location Address"
        dbInteger "ColumnWidth" ="4920"
        dbInteger "ColumnOrder" ="7"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BuyOut Cost"
        dbInteger "ColumnWidth" ="1830"
        dbInteger "ColumnOrder" ="8"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="0"
    End
    Begin
        dbText "Name" ="vw_Sch_Grp_Units_Remarketing.status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.BuyOutDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InterimDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AccountNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InterimBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InterimAmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InterimPaidDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="FAAmount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Buyer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FAPercent"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ActualSoldPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Final Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Adjustment"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ActualOffLease"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DSPN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InspPassed"
        dbLong "AggregateType" ="-1"
    End
End
