Operation =1
Option =0
Begin InputTables
    Name ="MstrLease"
    Name ="SchGrp"
    Name ="Schedule"
End
Begin OutputColumns
    Alias ="ClientID"
    Expression ="MstrLease.Client"
    Alias ="MLID"
    Expression ="Schedule.SchMLRef"
    Alias ="ScheduleID"
    Expression ="SchGrp.SGrpSchID"
    Alias ="GroupId"
    Expression ="SchGrp.SGrpID"
    Alias ="FAID"
    Expression ="SchGrp.SGrpFAID"
    Alias ="BLCD"
    Expression ="SchGrp.BLCD"
    Alias ="DateGroupAdded"
    Expression ="Min(SchGrp.GroupAddedDate)"
    Alias ="DateSchAdded"
    Expression ="Min(Schedule.DateSchAdded)"
    Alias ="Group"
    Expression ="SchGrp.SGrpNO"
    Alias ="Schedule"
    Expression ="Schedule.SchNo"
    Expression ="Schedule.ApprovalRequest"
    Expression ="Schedule.ApprovalReply"
End
Begin Joins
    LeftTable ="SchGrp"
    RightTable ="Schedule"
    Expression ="SchGrp.SGrpSchID = Schedule.SchID"
    Flag =1
    LeftTable ="MstrLease"
    RightTable ="Schedule"
    Expression ="MstrLease.MLID = Schedule.SchMLRef"
    Flag =1
End
Begin Groups
    Expression ="MstrLease.Client"
    GroupLevel =0
    Expression ="Schedule.SchMLRef"
    GroupLevel =0
    Expression ="SchGrp.SGrpSchID"
    GroupLevel =0
    Expression ="SchGrp.SGrpID"
    GroupLevel =0
    Expression ="SchGrp.SGrpFAID"
    GroupLevel =0
    Expression ="SchGrp.BLCD"
    GroupLevel =0
    Expression ="SchGrp.SGrpNO"
    GroupLevel =0
    Expression ="Schedule.SchNo"
    GroupLevel =0
    Expression ="Schedule.ApprovalRequest"
    GroupLevel =0
    Expression ="Schedule.ApprovalReply"
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
        dbText "Name" ="MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DateGroupAdded"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClientId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DateSchAdded"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.ApprovalReply"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.ApprovalRequest"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =83
    Top =-1
    Right =1287
    Bottom =988
    Left =-1
    Top =-1
    Right =1172
    Bottom =621
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =763
        Top =41
        Right =907
        Bottom =185
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =120
        Top =53
        Right =351
        Bottom =680
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =529
        Top =65
        Right =745
        Bottom =403
        Top =0
        Name ="Schedule"
        Name =""
    End
End
