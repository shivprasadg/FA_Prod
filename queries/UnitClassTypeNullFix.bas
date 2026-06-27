Operation =4
Option =0
Where ="((([Units].[UnitClassType]) Is Null))"
Begin InputTables
    Name ="Units"
End
Begin OutputColumns
    Alias ="Expr1"
    Name ="Units.UnitClassType"
    Expression ="IIf(\"UnitSubType\"=\"Tractor\",\"Tractor\",\"Trailer\")"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitClassType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitSubType"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =238
    Top =13
    Right =1720
    Bottom =1011
    Left =-1
    Top =-1
    Right =1450
    Bottom =697
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
End
