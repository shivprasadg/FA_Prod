Operation =1
Option =2
Where ="(((Units.UnitStatus)<>\"A\" And (Units.UnitStatus)<>\"B\"))"
Begin InputTables
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="Rmkt"
End
Begin OutputColumns
    Expression ="Schedule.SchID"
    Expression ="Schedule.SchNo"
    Expression ="MstrLease.MLID"
    Alias ="Units"
    Expression ="Count(Units.UnitID)"
    Expression ="MstrLease.Client"
    Expression ="MstrLease.MLNo"
    Expression ="MstrLease.MLOrig"
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
    LeftTable ="MstrLease"
    RightTable ="Schedule"
    Expression ="MstrLease.MLID = Schedule.SchMLRef"
    Flag =1
    LeftTable ="Units"
    RightTable ="Rmkt"
    Expression ="Units.UnitID = Rmkt.UnitRef"
    Flag =1
End
Begin OrderBy
    Expression ="Schedule.SchNo"
    Flag =0
    Expression ="MstrLease.MLID"
    Flag =0
End
Begin Groups
    Expression ="Schedule.SchID"
    GroupLevel =0
    Expression ="Schedule.SchNo"
    GroupLevel =0
    Expression ="MstrLease.MLID"
    GroupLevel =0
    Expression ="MstrLease.Client"
    GroupLevel =0
    Expression ="MstrLease.MLNo"
    GroupLevel =0
    Expression ="MstrLease.MLOrig"
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
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.Client"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLOrig"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-256
    Top =100
    Right =1286
    Bottom =817
    Left =-1
    Top =-1
    Right =1518
    Bottom =236
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =16
        Top =21
        Right =160
        Bottom =165
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =196
        Top =87
        Right =340
        Bottom =231
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =373
        Top =127
        Right =517
        Bottom =271
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =557
        Top =169
        Right =701
        Bottom =313
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =733
        Top =196
        Right =919
        Bottom =320
        Top =0
        Name ="Rmkt"
        Name =""
    End
End
