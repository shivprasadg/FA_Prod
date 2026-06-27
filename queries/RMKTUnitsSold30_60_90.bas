Operation =1
Option =0
Where ="(((Rmkt.InvoicePaidDate) Is Not Null))"
Begin InputTables
    Name ="Rmkt"
End
Begin OutputColumns
    Expression ="Rmkt.RmkSm"
    Alias ="UnitsSold"
    Expression ="Count(Rmkt.UnitRef)"
    Alias ="UnitsSoldLast30"
    Expression ="Sum(IIf([InvoicePaidDate] Between Date() And Date()-31,1,0))"
    Alias ="UnitsSoldLast60"
    Expression ="Sum(IIf([InvoicePaidDate] Between Date()-31 And Date()-62,1,0))"
    Alias ="UnitsSoldLast90"
    Expression ="Sum(IIf([InvoicePaidDate] Between Date()-63 And Date()-93,1,0))"
    Alias ="AmtSold30"
    Expression ="Sum(IIf([InvoicePaidDate] Between Date() And Date()-31,[SoldPrice],0))"
    Alias ="AmtSold60"
    Expression ="Sum(IIf([InvoicePaidDate] Between Date()-31 And Date()-62,[SoldPrice],0))"
    Alias ="AmtSold90"
    Expression ="Sum(IIf([InvoicePaidDate] Between Date()-63 And Date()-93,[SoldPrice],0))"
End
Begin Groups
    Expression ="Rmkt.RmkSm"
    GroupLevel =0
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
        dbText "Name" ="UnitsSold"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsSoldLast30"
        dbInteger "ColumnWidth" ="1875"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsSoldLast60"
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
        dbText "Name" ="AmtSold60"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AmtSold90"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AmtSold30"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =44
    Top =93
    Right =1482
    Bottom =1021
    Left =-1
    Top =-1
    Right =1406
    Bottom =610
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =620
        Bottom =460
        Top =0
        Name ="Rmkt"
        Name =""
    End
End
