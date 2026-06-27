Operation =1
Option =8
Where ="(((Rmkt.InvoicePaidDate) Between [forms]![acctmenu]![StartDate] And [forms]![acc"
    "tmenu]![EndDate]))"
Begin InputTables
    Name ="vw_Sch_Grp_Units_Remarketing"
    Name ="Units"
    Name ="Locations"
    Name ="Rmkt"
    Name ="Banks"
    Name ="Clients"
End
Begin OutputColumns
    Alias ="Client"
    Expression ="vw_Sch_Grp_Units_Remarketing.ClientShortName"
    Expression ="vw_Sch_Grp_Units_Remarketing.AssetID"
    Expression ="vw_Sch_Grp_Units_Remarketing.UnitNum"
    Expression ="vw_Sch_Grp_Units_Remarketing.VIN"
    Alias ="Description"
    Expression ="[MYear] & \" \" & [Make] & \"  \" & [Model] & \" \" & [Type] & \" \" & [axle]"
    Alias ="Location"
    Expression ="Locations.LocationName"
    Alias ="Location Address"
    Expression ="[LAddrLine1] & \" \" & [LCity] & \" \" & [LState] & \" \" & [LZip]"
    Alias ="BuyOut Cost"
    Expression ="CCur(Nz([Rmkt].[BuyOutCost],0))"
    Expression ="Rmkt.InvoicePaidDate"
    Expression ="Rmkt.BuyOutDate"
    Expression ="vw_Sch_Grp_Units_Remarketing.Status"
    Alias ="InterimBank"
    Expression ="Banks.BankShortName"
    Alias ="InterimAmt"
    Expression ="CCur(Nz([Rmkt.IFAmount],0))"
    Alias ="AccountNumber"
    Expression ="Rmkt.IFLoanNumber"
    Alias ="InterimDate"
    Expression ="Rmkt.IFDate"
    Alias ="InterimPaidDate"
    Expression ="Rmkt.IFPaidDate"
    Alias ="FAAmount"
    Expression ="CCur(Nz([Rmkt].[BuyOutCost],0))-(Nz([Rmkt].[IFAmount],0))"
    Alias ="FAPercent"
    Expression ="IIf(IsNull([Rmkt].[IFAmount]),Null,Round(1-([Rmkt].[IFAmount]/[Rmkt].[BuyOutCost"
        "]),3))"
    Alias ="Buyer"
    Expression ="Clients.ClientShNm"
    Alias ="SoldPrice"
    Expression ="CCur(Nz([Rmkt].[SoldPrice],0))"
    Alias ="Reserve"
    Expression ="IIf(Nz([Rmkt].[SoldPrice],0)<Nz([Rmkt].[BuyOutCost],0),CCur(Nz([Rmkt].[SoldPrice"
        "],0)-Nz([Rmkt].[BuyOutCost],0)),0)"
    Alias ="Profit"
    Expression ="IIf(Nz([Rmkt].[SoldPrice],0)<Nz([Rmkt].[BuyOutCost],0),0,CCur(Nz([Rmkt].[SoldPri"
        "ce],0)-Nz([Rmkt].[BuyOutCost],0)))"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="Locations"
    Expression ="Units.UnitGarageLocationRef = Locations.LocationID"
    Flag =2
    LeftTable ="vw_Sch_Grp_Units_Remarketing"
    RightTable ="Units"
    Expression ="vw_Sch_Grp_Units_Remarketing.AssetID = Units.UnitID"
    Flag =1
    LeftTable ="Units"
    RightTable ="Rmkt"
    Expression ="Units.UnitID = Rmkt.UnitRef"
    Flag =1
    LeftTable ="Rmkt"
    RightTable ="Banks"
    Expression ="Rmkt.IFBankID = Banks.BankId"
    Flag =2
    LeftTable ="Rmkt"
    RightTable ="Clients"
    Expression ="Rmkt.BuyerRef = Clients.clientID"
    Flag =2
End
Begin OrderBy
    Expression ="vw_Sch_Grp_Units_Remarketing.ClientShortName"
    Flag =0
    Expression ="vw_Sch_Grp_Units_Remarketing.AssetID"
    Flag =0
    Expression ="vw_Sch_Grp_Units_Remarketing.UnitNum"
    Flag =0
End
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
        dbText "Name" ="SoldPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Reserve"
        dbLong "AggregateType" ="0"
        dbInteger "ColumnWidth" ="1800"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Profit"
        dbInteger "ColumnWidth" ="2070"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="0"
    End
End
Begin
    State =0
    Left =293
    Top =128
    Right =2642
    Bottom =988
    Left =-1
    Top =-1
    Right =2317
    Bottom =232
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =117
        Top =8
        Right =293
        Bottom =434
        Top =0
        Name ="vw_Sch_Grp_Units_Remarketing"
        Name =""
    End
    Begin
        Left =344
        Top =101
        Right =488
        Bottom =425
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =524
        Top =154
        Right =708
        Bottom =404
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =737
        Top =42
        Right =985
        Bottom =330
        Top =0
        Name ="Rmkt"
        Name =""
    End
    Begin
        Left =1037
        Top =12
        Right =1188
        Bottom =235
        Top =0
        Name ="Banks"
        Name =""
    End
    Begin
        Left =1236
        Top =12
        Right =1380
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
End
