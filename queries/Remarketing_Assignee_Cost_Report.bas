Operation =1
Option =0
Where ="(((RemarketingReportAllUnits.SourceCompany)<>\"\"))"
Begin InputTables
    Name ="UnitOEC2"
    Name ="Units"
    Name ="Locations"
    Name ="SchGrp"
    Name ="RemarketReportExpenesTotal"
    Name ="RemarketingReportAllUnits"
End
Begin OutputColumns
    Expression ="RemarketingReportAllUnits.SourceCompany"
    Expression ="RemarketingReportAllUnits.Schedule"
    Alias ="Group"
    Expression ="RemarketingReportAllUnits.UnitGroup"
    Alias ="Unit#"
    Expression ="RemarketingReportAllUnits.Unitnum"
    Expression ="RemarketingReportAllUnits.UnitVIN"
    Expression ="RemarketingReportAllUnits.UnitStatus"
    Expression ="RemarketingReportAllUnits.UnitYr"
    Expression ="RemarketingReportAllUnits.UnitMake"
    Expression ="RemarketingReportAllUnits.UnitModel"
    Expression ="RemarketingReportAllUnits.UnitType"
    Expression ="RemarketingReportAllUnits.Location"
    Alias ="Original Cost"
    Expression ="UnitOEC2.TotalCost"
    Alias ="Monthly Payment"
    Expression ="Units.UnitRent"
    Alias ="BLCD"
    Expression ="SchGrp.BLCD"
    Alias ="Obligated Term"
    Expression ="CInt([LeaseTermPrimary])"
    Alias ="Expiration Date"
    Expression ="DateAdd('m',CInt([LeaseTermPrimary]),[BLCD])-1"
    Alias ="Parked Date"
    Expression ="RemarketingReportAllUnits.ParkDate"
    Alias ="Off Lease Date"
    Expression ="Units.unitoffleasedt"
    Expression ="RemarketingReportAllUnits.BuyOutDate"
    Alias ="Assignee Bank"
    Expression ="RemarketingReportAllUnits.BankPartner"
    Alias ="Parked Odometer"
    Expression ="RemarketingReportAllUnits.Odometer"
    Alias ="Residual"
    Expression ="Units.UnitSynResid"
    Alias ="Buy Out Cost"
    Expression ="RemarketingReportAllUnits.BuyOut"
    Alias ="Actual Sold Price"
    Expression ="RemarketingReportAllUnits.UnitSalePrice"
    Alias ="Remarketing Expenses"
    Expression ="RemarketingReportAllUnits.TotalExpense"
    Alias ="Gain/Loss On Sale"
    Expression ="RemarketingReportAllUnits.TotalProfit"
    Expression ="RemarketingReportAllUnits.ClientID"
    Expression ="RemarketingReportAllUnits.UnitID"
    Expression ="RemarketingReportAllUnits.InvoicePaid"
    Expression ="RemarketingReportAllUnits.InvoiceSent"
    Expression ="RemarketingReportAllUnits.BuyOutDate"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="RemarketingReportAllUnits"
    Expression ="Units.UnitID = RemarketingReportAllUnits.UnitID"
    Flag =1
    LeftTable ="Locations"
    RightTable ="RemarketingReportAllUnits"
    Expression ="Locations.LocationID = RemarketingReportAllUnits.LocationID"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="RemarketingReportAllUnits"
    Expression ="SchGrp.SGrpID = RemarketingReportAllUnits.SGrpID"
    Flag =1
    LeftTable ="RemarketingReportAllUnits"
    RightTable ="UnitOEC2"
    Expression ="RemarketingReportAllUnits.UnitID = UnitOEC2.AssetId"
    Flag =1
    LeftTable ="RemarketingReportAllUnits"
    RightTable ="RemarketReportExpenesTotal"
    Expression ="RemarketingReportAllUnits.UnitID = RemarketReportExpenesTotal.AssetId"
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
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.VIN"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Schedule"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Location"
        dbInteger "ColumnWidth" ="2685"
        dbInteger "ColumnOrder" ="11"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Monthly Payment"
        dbInteger "ColumnOrder" ="13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Residual"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="22"
    End
    Begin
        dbText "Name" ="Parked Odometer"
        dbInteger "ColumnOrder" ="19"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Actual Sold Price"
        dbInteger "ColumnOrder" ="25"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Parked Date"
        dbInteger "ColumnOrder" ="17"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Buy Out Cost"
        dbInteger "ColumnOrder" ="23"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Off Lease Date"
        dbInteger "ColumnOrder" ="18"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Assignee Bank"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="20"
    End
    Begin
        dbText "Name" ="BLCD"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="Original Cost"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="Gain/Loss On Sale"
        dbInteger "ColumnWidth" ="2070"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Remarketing Expenses"
        dbInteger "ColumnOrder" ="24"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expiration Date"
        dbInteger "ColumnOrder" ="16"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Obligated Term"
        dbInteger "ColumnOrder" ="15"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Group"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unit#"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitType"
        dbInteger "ColumnWidth" ="2685"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitYr"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitMake"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitStatus"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitModel"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sourceCompany"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BuyOutDate"
        dbInteger "ColumnWidth" ="2385"
        dbInteger "ColumnOrder" ="21"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.Schedule"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitYr"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.Location"
        dbInteger "ColumnWidth" ="2685"
        dbInteger "ColumnOrder" ="11"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.InvoicePaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.VIN"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitMake"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.ParkDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.InvoiceSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitVIN"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitModel"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.BuyOutDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.TotalExpense"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.Odometer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitSalePrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.TotalProfit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.BuyOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitStatus"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitType"
        dbInteger "ColumnWidth" ="2685"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.BankPartner"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitoffleasedt"
        dbInteger "ColumnOrder" ="18"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.SourceCompany"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-31
    Top =8
    Right =1695
    Bottom =889
    Left =-1
    Top =-1
    Right =1694
    Bottom =660
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =1461
        Top =35
        Right =1605
        Bottom =378
        Top =0
        Name ="UnitOEC2"
        Name =""
    End
    Begin
        Left =922
        Top =101
        Right =1071
        Bottom =577
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =576
        Top =146
        Right =720
        Bottom =428
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =572
        Top =466
        Right =801
        Bottom =682
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =1120
        Top =414
        Right =1485
        Bottom =639
        Top =0
        Name ="RemarketReportExpenesTotal"
        Name =""
    End
    Begin
        Left =231
        Top =31
        Right =449
        Bottom =463
        Top =0
        Name ="RemarketingReportAllUnits"
        Name =""
    End
End
