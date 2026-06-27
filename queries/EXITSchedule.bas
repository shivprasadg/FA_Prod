Operation =1
Option =0
Begin InputTables
    Name ="SchGrp"
    Name ="Units"
    Name ="Schedule"
End
Begin OutputColumns
    Expression ="Schedule.SchID"
    Alias ="MLID"
    Expression ="Schedule.SchMLRef"
    Alias ="Schedule"
    Expression ="Schedule.SchNo"
    Alias ="Units"
    Expression ="Count(Units.UnitID)"
End
Begin Joins
    LeftTable ="SchGrp"
    RightTable ="Schedule"
    Expression ="SchGrp.SGrpSchID = Schedule.SchID"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
End
Begin OrderBy
    Expression ="Schedule.SchID"
    Flag =0
    Expression ="Schedule.SchNo"
    Flag =0
End
Begin Groups
    Expression ="Schedule.SchID"
    GroupLevel =0
    Expression ="Schedule.SchMLRef"
    GroupLevel =0
    Expression ="Schedule.SchNo"
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
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MLID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =119
    Top =172
    Right =2270
    Bottom =842
    Left =-1
    Top =-1
    Right =2119
    Bottom =200
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =195
        Top =8
        Right =339
        Bottom =429
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =641
        Top =27
        Right =840
        Bottom =437
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =460
        Top =12
        Right =604
        Bottom =156
        Top =0
        Name ="Schedule"
        Name =""
    End
End
