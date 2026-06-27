Operation =4
Option =0
Where ="(((Units.UnitID)=[forms]![moveunit]![unit]))"
Begin InputTables
    Name ="Units"
End
Begin OutputColumns
    Name ="Units.GroupREF"
    Expression ="[forms]![moveunit]![newgroup]"
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
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.GroupREF"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =92
    Top =171
    Right =1240
    Bottom =786
    Left =-1
    Top =-1
    Right =1116
    Bottom =293
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
End
