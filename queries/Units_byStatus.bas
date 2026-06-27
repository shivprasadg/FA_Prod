Operation =1
Option =0
Begin InputTables
    Name ="Units"
    Name ="UnitStatus"
End
Begin OutputColumns
    Expression ="Units.UnitStatus"
    Alias ="UnitCount"
    Expression ="Count(Units.UnitID)"
    Expression ="UnitStatus.DEF"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="UnitStatus"
    Expression ="Units.UnitStatus = UnitStatus.Unitstatus"
    Flag =1
End
Begin OrderBy
    Expression ="Units.UnitStatus"
    Flag =0
End
Begin Groups
    Expression ="Units.UnitStatus"
    GroupLevel =0
    Expression ="UnitStatus.DEF"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbMemo "Filter" ="([Query7].[DEF] Like \"*owned*\")"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="-1"
Begin
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitStatus.DEF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitCount"
        dbLong "AggregateType" ="0"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1612
    Bottom =963
    Left =-1
    Top =-1
    Right =1580
    Bottom =623
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =76
        Top =18
        Right =278
        Bottom =369
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =298
        Top =54
        Right =442
        Bottom =198
        Top =0
        Name ="UnitStatus"
        Name =""
    End
End
