Operation =1
Option =0
Begin InputTables
    Name ="Units"
End
Begin OutputColumns
    Expression ="Units.GroupREF"
    Alias ="CountOfUnitID"
    Expression ="Count(Units.UnitID)"
End
Begin Groups
    Expression ="Units.GroupREF"
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
        dbText "Name" ="Units.GroupREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfUnitID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1449
    Bottom =852
    Left =-1
    Top =-1
    Right =1417
    Bottom =370
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =319
        Bottom =403
        Top =0
        Name ="Units"
        Name =""
    End
End
