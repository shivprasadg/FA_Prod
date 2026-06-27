Operation =1
Option =0
Where ="(((Units.UnitAcceptDate)>[UnitInServ]))"
Begin InputTables
    Name ="Units"
End
Begin OutputColumns
    Expression ="Units.UnitID"
    Expression ="Units.UnitVIN"
    Expression ="Units.UnitInServ"
    Expression ="Units.UnitAcceptDate"
    Expression ="Units.UnitDocAccpt"
    Alias ="DaysDiff"
    Expression ="[UnitAcceptDate]-[UnitInServ]"
    Expression ="Units.UnitYr"
    Expression ="Units.UnitMake"
    Expression ="Units.UnitModel"
    Expression ="Units.UnitType"
End
Begin OrderBy
    Expression ="Units.UnitInServ"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbMemo "OrderBy" ="[AcceptedDate_vs_InserviceDate_part1].[DaysDiff] DESC"
Begin
    Begin
        dbText "Name" ="Units.UnitInServ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitAcceptDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDocAccpt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbInteger "ColumnWidth" ="2205"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DaysDiff"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitYr"
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
        dbText "Name" ="Units.UnitType"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1520
    Bottom =825
    Left =-1
    Top =-1
    Right =1488
    Bottom =474
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =344
        Top =50
        Right =612
        Bottom =458
        Top =0
        Name ="Units"
        Name =""
    End
End
