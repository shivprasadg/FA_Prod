Operation =4
Option =0
Where ="(((SchGrp.SGrpID)=[forms]![movegrp]![group]))"
Begin InputTables
    Name ="SchGrp"
End
Begin OutputColumns
    Name ="SchGrp.SGrpSchID"
    Expression ="[forms]![movegrp]![newsch]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Groups.groupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpSchID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-42
    Top =-8
    Right =659
    Bottom =653
    Left =-1
    Top =-1
    Right =669
    Bottom =344
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="SchGrp"
        Name =""
    End
End
