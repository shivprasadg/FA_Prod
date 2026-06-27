Operation =1
Option =0
Begin InputTables
    Name ="RemarketingReportAllUnits"
End
Begin OutputColumns
    Alias ="Account Name"
    Expression ="RemarketingReportAllUnits.SourceCompany"
    Alias ="Lease Sch"
    Expression ="RemarketingReportAllUnits.Schedule"
    Alias ="Group"
    Expression ="RemarketingReportAllUnits.UnitGroup"
    Alias ="Customer Unit#"
    Expression ="RemarketingReportAllUnits.UnitNum"
    Alias ="FA Unit Status"
    Expression ="RemarketingReportAllUnits.UnitStatus"
    Alias ="Year"
    Expression ="RemarketingReportAllUnits.UnitYr"
    Alias ="Make"
    Expression ="RemarketingReportAllUnits.UnitMake"
    Alias ="Model"
    Expression ="RemarketingReportAllUnits.UnitModel"
    Alias ="Type"
    Expression ="RemarketingReportAllUnits.UnitType"
    Expression ="RemarketingReportAllUnits.Odometer"
    Expression ="RemarketingReportAllUnits.Buyer"
    Alias ="Disposition"
    Expression ="RemarketingReportAllUnits.BuyerType"
    Alias ="Target Sale Price"
    Expression ="RemarketingReportAllUnits.TargetSalePrice"
    Alias ="Sale Price"
    Expression ="RemarketingReportAllUnits.UnitSalePrice"
    Alias ="Days Since BuyOut"
    Expression ="RemarketingReportAllUnits.DaysInInventoryPerAccounting"
    Expression ="RemarketingReportAllUnits.BuyOutDate"
    Expression ="RemarketingReportAllUnits.UnitID"
    Expression ="RemarketingReportAllUnits.ClientID"
    Expression ="RemarketingReportAllUnits.FMV"
    Expression ="RemarketingReportAllUnits.GainLoss"
    Expression ="RemarketingReportAllUnits.TargetSalePrice"
    Expression ="RemarketingReportAllUnits.LocationID"
    Expression ="RemarketingReportAllUnits.PCity"
    Expression ="RemarketingReportAllUnits.PState"
    Expression ="RemarketingReportAllUnits.BuyerID"
    Expression ="RemarketingReportAllUnits.UnitYr"
    Expression ="RemarketingReportAllUnits.UnitMake"
    Expression ="RemarketingReportAllUnits.UnitModel"
    Expression ="RemarketingReportAllUnits.UnitType"
    Expression ="RemarketingReportAllUnits.UnitStatus"
    Expression ="RemarketingReportAllUnits.TitleRecv"
    Expression ="RemarketingReportAllUnits.InvoiceSent"
    Expression ="RemarketingReportAllUnits.[SUF Returned]"
    Expression ="RemarketingReportAllUnits.InvoicePaid"
    Expression ="RemarketingReportAllUnits.SalesPerson"
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
        dbText "Name" ="RemarketingReportAllUnits.UnitYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.LocationState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.Odometer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.LocationCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.TargetSalePrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Account Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.ClientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.FMV"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.BuyerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Year"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="825"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Make"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Model"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Target Sale Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sale Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Disposition"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.Buyer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lease Sch"
        dbInteger "ColumnWidth" ="3255"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FA Unit Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Customer Unit#"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.TitleRecv"
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
        dbText "Name" ="RemarketingReportAllUnits.BuyOutDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.[SUF Returned]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Days Since BuyOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.SalesPerson"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.PCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.PState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.GainLoss"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =179
    Top =266
    Right =1677
    Bottom =839
    Left =-1
    Top =-1
    Right =1474
    Bottom =178
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =71
        Top =12
        Right =372
        Bottom =211
        Top =0
        Name ="RemarketingReportAllUnits"
        Name =""
    End
End
