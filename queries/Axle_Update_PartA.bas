Operation =1
Option =0
Where ="(((Units.Axle)=\"TBD\"))"
Begin InputTables
    Name ="Units"
End
Begin OutputColumns
    Expression ="Units.UnitID"
    Expression ="Units.UnitVIN"
    Expression ="Units.Axle"
    Expression ="Units.UnitMake"
    Expression ="Units.UnitModel"
    Expression ="Units.UnitYr"
    Alias ="VinPrefix"
    Expression ="Left([UnitVIN],11)"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.Axle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbInteger "ColumnWidth" ="2265"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VinPrefix"
        dbInteger "ColumnWidth" ="2265"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =163
    Top =65
    Right =1683
    Bottom =850
    Left =-1
    Top =-1
    Right =1488
    Bottom =451
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =45
        Top =87
        Right =319
        Bottom =379
        Top =0
        Name ="Units"
        Name =""
    End
End
