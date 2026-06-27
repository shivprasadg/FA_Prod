Operation =1
Option =0
Begin InputTables
    Name ="Bills"
End
Begin OutputColumns
    Alias ="BilledSchID"
    Expression ="Bills.BillSchRef"
    Expression ="Bills.BillType"
    Alias ="BilledFrom"
    Expression ="Min(Bills.Billfrom)"
    Alias ="BilledTo"
    Expression ="Max(Bills.Billto)"
End
Begin Groups
    Expression ="Bills.BillSchRef"
    GroupLevel =0
    Expression ="Bills.BillType"
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
        dbText "Name" ="Bills.BillType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BilledFrom"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BilledSchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BilledTo"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =47
    Top =154
    Right =1520
    Bottom =827
    Left =-1
    Top =-1
    Right =1441
    Bottom =293
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =256
        Bottom =311
        Top =0
        Name ="Bills"
        Name =""
    End
End
