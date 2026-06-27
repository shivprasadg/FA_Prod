Operation =1
Option =0
Begin InputTables
    Name ="Units"
End
Begin OutputColumns
    Expression ="Units.UnitID"
    Expression ="Units.EstDelInd"
    Expression ="Units.UnitOriginserv"
    Expression ="Units.UnitDep"
    Expression ="Units.UnitDepdt"
    Expression ="Units.UnitDepPdDT"
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
        dbText "Name" ="Units.UnitDep"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDepdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDepPdDT"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1650"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Units.UnitOriginserv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.EstDelInd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =5
    Top =6
    Right =1549
    Bottom =1000
    Left =-1
    Top =-1
    Right =1526
    Bottom =700
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =77
        Top =3
        Right =281
        Bottom =707
        Top =0
        Name ="Units"
        Name =""
    End
End
