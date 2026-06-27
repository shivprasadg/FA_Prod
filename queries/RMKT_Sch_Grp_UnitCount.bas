Operation =1
Option =0
Having ="(((Schedule.SchFAID)<>\"\") AND ((Units.UnitStatus)=\"TR\" Or (Units.UnitStatus)"
    "=\"R\" Or (Units.UnitStatus)=\"S\"))"
Begin InputTables
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
End
Begin OutputColumns
    Alias ="FAID"
    Expression ="Schedule.SchFAID"
    Alias ="Schedule"
    Expression ="Schedule.SchNo"
    Alias ="Group"
    Expression ="SchGrp.SGrpNO"
    Alias ="Units"
    Expression ="Count(Units.UnitID)"
    Expression ="SchGrp.SGrpID"
    Expression ="Schedule.SchID"
    Expression ="Units.UnitStatus"
End
Begin Joins
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
End
Begin OrderBy
    Expression ="Schedule.SchNo"
    Flag =0
    Expression ="SchGrp.SGrpNO"
    Flag =0
End
Begin Groups
    Expression ="Schedule.SchFAID"
    GroupLevel =0
    Expression ="Schedule.SchNo"
    GroupLevel =0
    Expression ="SchGrp.SGrpNO"
    GroupLevel =0
    Expression ="SchGrp.SGrpID"
    GroupLevel =0
    Expression ="Schedule.SchID"
    GroupLevel =0
    Expression ="Units.UnitStatus"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="FAID"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UNITS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Group"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =12
    Top =38
    Right =1592
    Bottom =898
    Left =-1
    Top =-1
    Right =1548
    Bottom =475
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =220
        Bottom =383
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =381
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =407
        Top =12
        Right =654
        Bottom =464
        Top =0
        Name ="Units"
        Name =""
    End
End
