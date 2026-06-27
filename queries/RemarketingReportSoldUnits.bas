Operation =1
Option =0
Where ="(((RemarketingReportAllUnits.InvoicePaid) Is Not Null))"
Begin InputTables
    Name ="RemarketingReportAllUnits"
End
Begin OutputColumns
    Alias ="Account Name"
    Expression ="RemarketingReportAllUnits.SourceCompany"
    Expression ="RemarketingReportAllUnits.Schedule"
    Expression ="RemarketingReportAllUnits.UnitGroup"
    Alias ="Unit#"
    Expression ="RemarketingReportAllUnits.UnitNum"
    Expression ="RemarketingReportAllUnits.UnitStatus"
    Expression ="RemarketingReportAllUnits.UnitYr"
    Expression ="RemarketingReportAllUnits.UnitMake"
    Expression ="RemarketingReportAllUnits.UnitModel"
    Expression ="RemarketingReportAllUnits.UnitType"
    Expression ="RemarketingReportAllUnits.Odometer"
    Expression ="RemarketingReportAllUnits.Buyer"
    Expression ="RemarketingReportAllUnits.BuyersOrder"
    Expression ="RemarketingReportAllUnits.DaysInInventory"
    Expression ="RemarketingReportAllUnits.Residual"
    Alias ="Term Penalty"
    Expression ="RemarketingReportAllUnits.TermPenalty"
    Alias ="BookValue"
    Expression ="[BuyOut]+Nz([TermPenalty],0)"
    Expression ="RemarketingReportAllUnits.BuyOut"
    Alias ="RMKT Expense"
    Expression ="RemarketingReportAllUnits.TotalExpense"
    Alias ="TotalCost"
    Expression ="([BuyOut]+Nz([TermPenalty],0))+[TotalExpense]"
    Alias ="Sale Price"
    Expression ="RemarketingReportAllUnits.UnitSalePrice"
    Alias ="PnL on Sale"
    Expression ="[UnitSalePrice]-(([BuyOut]+Nz([TermPenalty],0))+[TotalExpense])"
    Expression ="RemarketingReportAllUnits.OEC"
    Expression ="RemarketingReportAllUnits.InvoiceSent"
    Expression ="RemarketingReportAllUnits.InvoicePaid"
    Expression ="RemarketingReportAllUnits.InvoiceNum"
    Expression ="RemarketingReportAllUnits.SalesPerson"
    Expression ="RemarketingReportAllUnits.UnitVIN"
    Expression ="RemarketingReportAllUnits.UnitID"
    Alias ="Days Since BuyOut"
    Expression ="RemarketingReportAllUnits.DaysInInventoryPerAccounting"
    Expression ="RemarketingReportAllUnits.BuyOutDate"
    Expression ="RemarketingReportAllUnits.ClientID"
    Expression ="RemarketingReportAllUnits.BuyerID"
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
        dbText "Name" ="RemarketingReportAllUnits.BuyOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sale Price"
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
        dbText "Name" ="PnL on Sale"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKT Expense"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.Buyer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Term Penalty"
        dbLong "AggregateType" ="-1"
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
        dbText "Name" ="RemarketingReportAllUnits.InvoicePaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.DaysInInventory"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unit#"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Account Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.SalesPerson"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.unitvin"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.INVOICESENT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.buyoutdate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Days Since BuyOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BookValue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.Residual"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.OEC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.BuyersOrder"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.InvoiceNum"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =2
    Top =106
    Right =1639
    Bottom =1057
    Left =-1
    Top =-1
    Right =1613
    Bottom =325
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =440
        Top =-12
        Right =675
        Bottom =446
        Top =0
        Name ="RemarketingReportAllUnits"
        Name =""
    End
End
