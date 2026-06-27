Operation =1
Option =0
Where ="(((Units.UnitStatus)<>\"T\"))"
Begin InputTables
    Name ="Units"
End
Begin OutputColumns
    Expression ="Units.UnitType"
    Expression ="Units.UnitMake"
    Expression ="Units.UnitModel"
    Alias ="Units"
    Expression ="Count(Units.UnitID)"
End
Begin OrderBy
    Expression ="Units.UnitType"
    Flag =0
    Expression ="Units.UnitMake"
    Flag =0
    Expression ="Units.UnitModel"
    Flag =0
End
Begin Groups
    Expression ="Units.UnitType"
    GroupLevel =0
    Expression ="Units.UnitMake"
    GroupLevel =0
    Expression ="Units.UnitModel"
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
        dbText "Name" ="Units.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =58
    Top =74
    Right =1596
    Bottom =803
    Left =-1
    Top =-1
    Right =1506
    Bottom =421
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =78
        Top =123
        Right =474
        Bottom =518
        Top =0
        Name ="Units"
        Name =""
    End
End
