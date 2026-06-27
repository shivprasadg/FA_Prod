Operation =1
Option =0
Begin InputTables
    Name ="Mileage"
End
Begin OutputColumns
    Expression ="Mileage.UnitID"
    Alias ="MaxOfMHDate"
    Expression ="Max(Mileage.MHDate)"
End
Begin OrderBy
    Expression ="Mileage.UnitID"
    Flag =0
End
Begin Groups
    Expression ="Mileage.UnitID"
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
        dbText "Name" ="Mileage.UnitRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfMHDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =76
    Top =132
    Right =1544
    Bottom =793
    Left =-1
    Top =-1
    Right =1436
    Bottom =344
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Mileage"
        Name =""
    End
End
