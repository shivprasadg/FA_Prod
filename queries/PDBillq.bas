Operation =1
Option =0
Having ="(((Bills.BillType)=\"PerDiem\"))"
Begin InputTables
    Name ="Bills"
End
Begin OutputColumns
    Expression ="Bills.BillSchRef"
    Alias ="MaxOfBillSent"
    Expression ="Max(Bills.BillSent)"
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
        dbText "Name" ="Bills.BillSchRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfBillSent"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1258
    Bottom =686
    Left =-1
    Top =-1
    Right =1226
    Bottom =277
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Bills"
        Name =""
    End
End
