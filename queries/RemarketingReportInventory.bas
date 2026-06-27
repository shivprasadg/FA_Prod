Operation =1
Option =0
Begin InputTables
    Name ="RemarketingReportAllUnits"
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Alias ="Account Name"
    Expression ="RemarketingReportAllUnits.SourceCompany"
    Expression ="vw_SixKeys.Schedule"
    Expression ="vw_SixKeys.UnitGroup"
    Alias ="Cust Unit#"
    Expression ="vw_SixKeys.Unitnum"
    Expression ="RemarketingReportAllUnits.UnitStatus"
    Expression ="RemarketingReportAllUnits.UnitYr"
    Expression ="RemarketingReportAllUnits.UnitMake"
    Expression ="RemarketingReportAllUnits.UnitModel"
    Expression ="RemarketingReportAllUnits.UnitType"
    Alias ="Location"
    Expression ="vw_SixKeys.Locationsn"
    Alias ="ProjectedParkedDate"
    Expression ="\"\""
    Expression ="RemarketingReportAllUnits.DaysInInventory"
    Expression ="RemarketingReportAllUnits.Odometer"
    Alias ="TotalCost"
    Expression ="[BuyOut]+[TotalExpense]"
    Alias ="Actual Sale Price"
    Expression ="RemarketingReportAllUnits.UnitSalePrice"
    Alias ="PnL on Actual"
    Expression ="IIf([UnitSalePrice]>0,[UnitProfit],0)"
    Alias ="FMVBasePrice"
    Expression ="RemarketingReportAllUnits.FMVBase"
    Alias ="PnL on FMVBase"
    Expression ="IIf([FMVBasePrice]>0,[FMVBase]-[TotalCost],0)"
    Expression ="RemarketingReportAllUnits.FMV"
    Alias ="PnL on FMV"
    Expression ="IIf([FMV]>0,[FMV]-[TotalCost],0)"
    Expression ="RemarketingReportAllUnits.TargetSalePrice"
    Alias ="PnL on Target"
    Expression ="IIf([TargetSalePrice]>0,[TargetSalePrice]-[TotalCost],0)"
    Alias ="Term Penalty"
    Expression ="IIf([FMVBase]>0,[FMVBase]-[BuyOut],0)"
    Expression ="RemarketingReportAllUnits.BankPartner"
    Expression ="RemarketingReportAllUnits.BankShareYN"
    Expression ="RemarketingReportAllUnits.BankShare"
    Expression ="vw_SixKeys.clientID"
    Expression ="RemarketingReportAllUnits.BuyerID"
    Expression ="RemarketingReportAllUnits.UnitID"
    Expression ="RemarketingReportAllUnits.BuyOutDate"
    Expression ="RemarketingReportAllUnits.InvoiceSent"
    Expression ="RemarketingReportAllUnits.InvoicePaid"
    Alias ="Days Since BuyOut"
    Expression ="RemarketingReportAllUnits.DaysInInventoryPerAccounting"
End
Begin Joins
    LeftTable ="RemarketingReportAllUnits"
    RightTable ="vw_SixKeys"
    Expression ="RemarketingReportAllUnits.UnitID = vw_SixKeys.AssetID"
    Flag =1
End
Begin OrderBy
    Expression ="RemarketingReportAllUnits.SourceCompany"
    Flag =0
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
        dbText "Name" ="Account Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cust Unit#"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.BuyerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.Odometer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Term Penalty"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1650"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="TotalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.DaysInInventory"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ProjectedParkedDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Location"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.Bankshare"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FMVBasePrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PnL on FMVBase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Actual Sale Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PnL on Actual"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PnL on Target"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.TargetSalePrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.BankshareYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.BuyOutDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.FMV"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PnL on FMV"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.BankPartner"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.InvoicePaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.InvoiceSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BuyOutDays"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1140
    Bottom =1021
    Left =-1
    Top =-1
    Right =1108
    Bottom =536
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =435
        Top =88
        Right =670
        Bottom =546
        Top =0
        Name ="RemarketingReportAllUnits"
        Name =""
    End
    Begin
        Left =741
        Top =88
        Right =947
        Bottom =467
        Top =0
        Name ="vw_Client_FAID_PO_Lease_Sch_Grp_Units"
        Name =""
    End
End
