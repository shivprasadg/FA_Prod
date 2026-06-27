Operation =1
Option =0
Having ="(((SchGrp.LeaseType)=\"FMV\"))"
Begin InputTables
    Name ="Schedule"
    Name ="MstrLease"
    Name ="SchGrp"
End
Begin OutputColumns
    Alias ="ClientID"
    Expression ="MstrLease.Client"
    Expression ="MstrLease.MLID"
    Alias ="FAID"
    Expression ="Nz([SgrpFAID],[SchFAID])"
    Expression ="Schedule.DateSchAdded"
    Alias ="ScheduleID"
    Expression ="SchGrp.SGrpSchID"
    Alias ="Schedule"
    Expression ="Schedule.SchNo"
    Alias ="Assignee"
    Expression ="SchGrp.Assignee_Bank"
    Alias ="FirstBLCD"
    Expression ="Min(SchGrp.BLCD)"
    Expression ="SchGrp.LeaseType"
End
Begin Joins
    LeftTable ="Schedule"
    RightTable ="MstrLease"
    Expression ="Schedule.SchMLRef = MstrLease.MLID"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
End
Begin OrderBy
    Expression ="SchGrp.SGrpSchID"
    Flag =0
End
Begin Groups
    Expression ="MstrLease.Client"
    GroupLevel =0
    Expression ="MstrLease.MLID"
    GroupLevel =0
    Expression ="Nz([SgrpFAID],[SchFAID])"
    GroupLevel =0
    Expression ="Schedule.DateSchAdded"
    GroupLevel =0
    Expression ="SchGrp.SGrpSchID"
    GroupLevel =0
    Expression ="Schedule.SchNo"
    GroupLevel =0
    Expression ="SchGrp.Assignee_Bank"
    GroupLevel =0
    Expression ="SchGrp.LeaseType"
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
        dbText "Name" ="Assignee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstBLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.DateSchAdded"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FAID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1309
    Bottom =974
    Left =-1
    Top =-1
    Right =1277
    Bottom =532
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =330
        Top =68
        Right =474
        Bottom =476
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =137
        Top =79
        Right =281
        Bottom =444
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =543
        Top =69
        Right =826
        Bottom =576
        Top =0
        Name ="SchGrp"
        Name =""
    End
End
