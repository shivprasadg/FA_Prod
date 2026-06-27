Operation =1
Option =0
Where ="(((vw_SixKeys.SGrpID) Is Null))"
Begin InputTables
    Name ="SchGrp"
    Name ="vw_SixKeys"
    Name ="Schedule"
    Name ="MstrLease"
    Name ="Clients"
    Name ="Units"
End
Begin OutputColumns
    Expression ="Clients.ClientShNm"
    Expression ="MstrLease.MLNo"
    Expression ="Schedule.SchNo"
    Expression ="SchGrp.SGrpNO"
    Expression ="MstrLease.MLOrig"
    Expression ="vw_SixKeys.SGrpID"
    Expression ="Clients.clientID"
    Expression ="Schedule.SchID"
    Expression ="SchGrp.SGrpID"
    Expression ="Units.UnitID"
    Expression ="Clients.clientID"
End
Begin Joins
    LeftTable ="SchGrp"
    RightTable ="vw_SixKeys"
    Expression ="SchGrp.SGrpID = vw_SixKeys.SGrpID"
    Flag =2
    LeftTable ="SchGrp"
    RightTable ="Schedule"
    Expression ="SchGrp.SGrpSchID = Schedule.SchID"
    Flag =2
    LeftTable ="Schedule"
    RightTable ="MstrLease"
    Expression ="Schedule.SchMLRef = MstrLease.MLID"
    Flag =2
    LeftTable ="MstrLease"
    RightTable ="Clients"
    Expression ="MstrLease.Client = Clients.ClientGroupId"
    Flag =2
    LeftTable ="Units"
    RightTable ="SchGrp"
    Expression ="Units.GroupREF = SchGrp.SGrpID"
    Flag =3
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
        dbText "Name" ="Schedule.SchNo"
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
        dbText "Name" ="vw_SixKeys.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-67
    Top =12
    Right =1362
    Bottom =981
    Left =-1
    Top =-1
    Right =1411
    Bottom =692
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =315
        Top =107
        Right =459
        Bottom =612
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =508
        Top =79
        Right =652
        Bottom =223
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =698
        Top =8
        Right =842
        Bottom =152
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =914
        Top =109
        Right =1058
        Bottom =253
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =59
        Top =271
        Right =203
        Bottom =415
        Top =0
        Name ="Units"
        Name =""
    End
End
