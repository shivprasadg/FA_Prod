Operation =1
Option =0
Begin InputTables
    Name ="Units"
    Name ="SchGrp"
    Name ="Schedule"
End
Begin OutputColumns
    Alias ="MLID"
    Expression ="Schedule.SchMLRef"
    Expression ="Schedule.SchID"
    Alias ="GrpID"
    Expression ="SchGrp.SGrpID"
    Alias ="UnitCount"
    Expression ="Count(Units.UnitID)"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="SchGrp"
    Expression ="Units.GroupREF = SchGrp.SGrpID"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Schedule"
    Expression ="SchGrp.SGrpSchID = Schedule.SchID"
    Flag =1
End
Begin OrderBy
    Expression ="Schedule.SchID"
    Flag =1
    Expression ="SchGrp.SGrpID"
    Flag =1
End
Begin Groups
    Expression ="Schedule.SchMLRef"
    GroupLevel =0
    Expression ="Schedule.SchID"
    GroupLevel =0
    Expression ="SchGrp.SGrpID"
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
dbMemo "Filter" ="([UnitsInGroupCount].[MLID]=247)"
Begin
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GrpID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =324
    Top =125
    Right =1612
    Bottom =1059
    Left =-1
    Top =-1
    Right =1256
    Bottom =580
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="Schedule"
        Name =""
    End
End
