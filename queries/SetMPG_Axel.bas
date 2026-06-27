Operation =4
Option =0
Where ="(((Units.clientgroupid)=151))"
Begin InputTables
    Name ="Units"
    Name ="SchGrp"
    Name ="Schedule"
    Name ="MstrLease"
    Name ="Clients"
End
Begin OutputColumns
    Name ="Units.MpgGoal"
    Expression ="6"
    Name ="Units.DrivingMpgGoal"
    Expression ="6"
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
    LeftTable ="Schedule"
    RightTable ="MstrLease"
    Expression ="Schedule.SchMLRef = MstrLease.MLID"
    Flag =1
    LeftTable ="MstrLease"
    RightTable ="Clients"
    Expression ="MstrLease.Client = Clients.clientID"
    Flag =1
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
    Left =125
    Top =95
    Right =1346
    Bottom =832
    Left =-1
    Top =-1
    Right =1189
    Bottom =302
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
    Begin
        Left =503
        Top =68
        Right =647
        Bottom =212
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =687
        Top =69
        Right =831
        Bottom =277
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =852
        Top =49
        Right =1026
        Bottom =294
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =1055
        Top =17
        Right =1212
        Bottom =164
        Top =0
        Name ="Clients"
        Name =""
    End
End
