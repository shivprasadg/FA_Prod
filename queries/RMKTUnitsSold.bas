Operation =1
Option =0
Where ="(((Rmkt.InvoicePaidDate) Is Not Null))"
Begin InputTables
    Name ="Rmkt"
    Name ="RMKTSalesPersons"
End
Begin OutputColumns
    Expression ="Rmkt.RmkSm"
    Alias ="AssetID"
    Expression ="Rmkt.UnitRef"
    Expression ="Rmkt.InvoicePaidDate"
    Expression ="Rmkt.SoldPrice"
    Alias ="SpID"
    Expression ="RMKTSalesPersons.ID"
End
Begin Joins
    LeftTable ="Rmkt"
    RightTable ="RMKTSalesPersons"
    Expression ="Rmkt.RmkSm = RMKTSalesPersons.SP"
    Flag =1
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
        dbText "Name" ="Rmkt.RmkSm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmktSalePr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.InvoicePaidDate"
        dbInteger "ColumnWidth" ="1875"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AmtSold30"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKTSalesPersons.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsSold"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsSoldLast60"
        dbInteger "ColumnWidth" ="1875"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AmtSold60"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AmtSold90"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsSoldLast30"
        dbInteger "ColumnWidth" ="1875"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsSoldLast90"
        dbInteger "ColumnWidth" ="1875"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SpID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =277
    Top =72
    Right =1715
    Bottom =1000
    Left =-1
    Top =-1
    Right =1406
    Bottom =542
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =275
        Top =43
        Right =847
        Bottom =491
        Top =0
        Name ="Rmkt"
        Name =""
    End
    Begin
        Left =40
        Top =62
        Right =184
        Bottom =416
        Top =0
        Name ="RMKTSalesPersons"
        Name =""
    End
End
