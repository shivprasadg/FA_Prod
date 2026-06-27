Operation =4
Option =0
Where ="(((Units.ClientGroupId)=[Which ClientGroupID]))"
Begin InputTables
    Name ="Units"
End
Begin OutputColumns
    Name ="Units.MpgGoal"
    Expression ="[MPGOverAll]"
    Name ="Units.DrivingMpgGoal"
    Expression ="[MPGDriving]"
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
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
Begin
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.GroupREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.MpgGoal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.DrivingMpgGoal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.Axle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitstatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.clientgroupid"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-8
    Top =29
    Right =1213
    Bottom =766
    Left =-1
    Top =-1
    Right =1189
    Bottom =234
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =163
        Top =54
        Right =456
        Bottom =312
        Top =0
        Name ="Units"
        Name =""
    End
End
