Operation =2
Name ="ReportStartEnd_SoldUnits"
Option =0
Where ="(((Units.ReportEndDate)=#12/31/2030#))"
Begin InputTables
    Name ="Units"
    Name ="ReMarketing_UnitSales"
End
Begin OutputColumns
    Expression ="Units.UnitID"
    Expression ="Units.UnitStatus"
    Expression ="Units.ReportStartDate"
    Expression ="Units.ReportEndDate"
    Alias ="OffLease"
    Expression ="Units.unitoffleasedt"
    Alias ="SoldDate"
    Expression ="ReMarketing_UnitSales.InvoicePaidDate"
    Alias ="SoldPrice"
    Expression ="ReMarketing_UnitSales.SoldPrice"
    Alias ="BOSent"
    Expression ="ReMarketing_UnitSales.RmktBOSSent"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="ReMarketing_UnitSales"
    Expression ="Units.UnitVIN = ReMarketing_UnitSales.unitvin"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="Units.ReportStartDate"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="Units.ReportEndDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="Units.unitoffleasedt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitOLInd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.RmktSaleDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.RmktSalePr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ReMarketing_SalesCust.PmtReceived"
        dbInteger "ColumnWidth" ="2640"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ReMarketing_SalesCust.TargetSale"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1009"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ReMarketing_SalesCust.RmktBOSSent"
        dbInteger "ColumnWidth" ="1695"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ReMarketing_SalesCust.RmktSalePr"
        dbInteger "ColumnWidth" ="1740"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BOSent"
        dbInteger "ColumnWidth" ="1695"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="SoldPrice"
        dbInteger "ColumnWidth" ="1740"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="Expr1000"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OffLease"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="ReMarketing_UnitSales.PmtReceived"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PaidDate"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SoldDate"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1196
    Bottom =806
    Left =-1
    Top =-1
    Right =1164
    Bottom =386
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =47
        Top =70
        Right =261
        Bottom =392
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =309
        Top =12
        Right =602
        Bottom =385
        Top =0
        Name ="ReMarketing_UnitSales"
        Name =""
    End
End
