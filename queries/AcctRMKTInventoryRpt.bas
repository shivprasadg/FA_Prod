Operation =1
Option =0
Begin InputTables
    Name ="AcctRMKTInventoryRpt_Part1"
    Name ="Remarketing_AvgSoldPrice"
    Name ="Remarketing_AvgSoldPrice_dated"
    Name ="Remarketing_AvgProjectedReserve"
End
Begin OutputColumns
    Expression ="AcctRMKTInventoryRpt_Part1.*"
    Expression ="Remarketing_AvgProjectedReserve.ProjectedSalePrice"
    Alias ="Reserve"
    Expression ="IIf(Nz([ProjectedSalePrice],0)>1 And Nz([ProjectedSalePrice],0)<Nz([Final Cost],"
        "0),CCur(Nz([ProjectedSalePrice],0)-Nz([Final Cost],0)),0)"
    Alias ="ProjectedProfit"
    Expression ="IIf(Nz([Remarketing_AvgSoldPrice_dated].[AvgSoldPrice],0)>Nz([Final Cost],0),CCu"
        "r(Nz([Remarketing_AvgSoldPrice_dated].[AvgSoldPrice],0)-Nz([Final Cost],0)),0)"
    Alias ="DatedAvgSalePrice"
    Expression ="Remarketing_AvgSoldPrice_dated.AvgSoldPrice"
    Alias ="90DaysAvgSoldPrice"
    Expression ="Remarketing_AvgSoldPrice.AvgSoldPrice"
    Alias ="ActualProfit"
    Expression ="IIf(Nz([ActualSoldPrice],0)=0,0,IIf(Nz([ActualSoldPrice],0)>Nz([Final Cost],0),C"
        "Cur(Nz([ActualSoldPrice],0)-Nz([Final Cost],0)),0))"
    Alias ="ActualLoss"
    Expression ="IIf(Nz([ActualSoldPrice],0)=0,0,IIf(Nz([ActualSoldPrice],0)<Nz([Final Cost],0),C"
        "Cur(Nz([ActualSoldPrice],0)-Nz([Final Cost],0)),0))"
End
Begin Joins
    LeftTable ="AcctRMKTInventoryRpt_Part1"
    RightTable ="Remarketing_AvgSoldPrice"
    Expression ="AcctRMKTInventoryRpt_Part1.Description = Remarketing_AvgSoldPrice.Description"
    Flag =2
    LeftTable ="AcctRMKTInventoryRpt_Part1"
    RightTable ="Remarketing_AvgSoldPrice_dated"
    Expression ="AcctRMKTInventoryRpt_Part1.Description = Remarketing_AvgSoldPrice_dated.Descript"
        "ion"
    Flag =2
    LeftTable ="AcctRMKTInventoryRpt_Part1"
    RightTable ="Remarketing_AvgProjectedReserve"
    Expression ="AcctRMKTInventoryRpt_Part1.Description = Remarketing_AvgProjectedReserve.Descrip"
        "tion"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="AcctRMKTInventoryRpt_Part1.vw_Sch_Grp_Units_Remarketing.VIN"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="Reserve"
        dbInteger "ColumnWidth" ="2235"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="0"
        dbInteger "ColumnOrder" ="21"
    End
    Begin
        dbText "Name" ="AcctRMKTInventoryRpt_Part1.Client"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="AcctRMKTInventoryRpt_Part1.Rmkt.BuyOutDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="AcctRMKTInventoryRpt_Part1.vw_Sch_Grp_Units_Remarketing.Status"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="AcctRMKTInventoryRpt_Part1.FAAmount"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="17"
    End
    Begin
        dbText "Name" ="AcctRMKTInventoryRpt_Part1.vw_Sch_Grp_Units_Remarketing.AssetID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="AcctRMKTInventoryRpt_Part1.Description"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
        dbInteger "ColumnWidth" ="4500"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="AcctRMKTInventoryRpt_Part1.Location"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="AcctRMKTInventoryRpt_Part1.BuyOut Cost"
        dbLong "AggregateType" ="0"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="AcctRMKTInventoryRpt_Part1.Rmkt.InvoicePaidDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="AcctRMKTInventoryRpt_Part1.InterimBank"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="AcctRMKTInventoryRpt_Part1.InterimAmt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="AcctRMKTInventoryRpt_Part1.AccountNumber"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="AcctRMKTInventoryRpt_Part1.InterimDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="15"
    End
    Begin
        dbText "Name" ="AcctRMKTInventoryRpt_Part1.InterimPaidDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="16"
    End
    Begin
        dbText "Name" ="AcctRMKTInventoryRpt_Part1.FAPercent"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="18"
    End
    Begin
        dbText "Name" ="AcctRMKTInventoryRpt_Part1.vw_Sch_Grp_Units_Remarketing.UnitNum"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="AcctRMKTInventoryRpt_Part1.Location Address"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="AcctRMKTInventoryRpt_Part1.Buyer"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="19"
    End
    Begin
        dbText "Name" ="Remarketing_AvgProjectedReserve.ProjectedSalePrice"
        dbInteger "ColumnOrder" ="20"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DatedAvgSalePrice"
        dbInteger "ColumnOrder" ="23"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2175"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="90DaysAvgSoldPrice"
        dbInteger "ColumnWidth" ="2340"
        dbInteger "ColumnOrder" ="24"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ProjectedProfit"
        dbInteger "ColumnOrder" ="22"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ActualProfit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ActualLoss"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AcctRMKTInventoryRpt_Part1.ActualSoldPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AcctRMKTInventoryRpt_Part1.Adjustment"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AcctRMKTInventoryRpt_Part1.Final Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AcctRMKTInventoryRpt_Part1.ActualOffLease"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AcctRMKTInventoryRpt_Part1.DSPN"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =19
    Top =68
    Right =1664
    Bottom =1015
    Left =-1
    Top =-1
    Right =1621
    Bottom =379
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =78
        Top =25
        Right =266
        Bottom =393
        Top =0
        Name ="AcctRMKTInventoryRpt_Part1"
        Name =""
    End
    Begin
        Left =514
        Top =340
        Right =721
        Bottom =484
        Top =0
        Name ="Remarketing_AvgSoldPrice"
        Name =""
    End
    Begin
        Left =579
        Top =168
        Right =790
        Bottom =312
        Top =0
        Name ="Remarketing_AvgSoldPrice_dated"
        Name =""
    End
    Begin
        Left =338
        Top =12
        Right =615
        Bottom =156
        Top =0
        Name ="Remarketing_AvgProjectedReserve"
        Name =""
    End
End
