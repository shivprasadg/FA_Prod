Operation =1
Option =0
Where ="(((SchGrp.BLCD) Is Null) AND ((MstrLease.MLOrig)=\"FA\"))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="AtfRevq1"
    Name ="SchGrp"
End
Begin OutputColumns
    Expression ="Clients.clientgroupID"
    Expression ="Clients.ClientShNm"
    Expression ="Schedule.SchNo"
    Expression ="SchGrp.BLCD"
    Expression ="AtfRevq1.RevNo"
    Expression ="AtfRevq1.RevDt"
    Expression ="AtfRevq1.ReviewDt"
    Expression ="AtfRevq1.reviewer"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="MstrLease"
    Expression ="Clients.clientID = MstrLease.Client"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="AtfRevq1"
    Expression ="Schedule.SchID = AtfRevq1.schID"
    Flag =1
    LeftTable ="MstrLease"
    RightTable ="Schedule"
    Expression ="MstrLease.MLID = Schedule.SchMLRef"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
End
Begin OrderBy
    Expression ="Clients.clientgroupID"
    Flag =0
    Expression ="Clients.ClientShNm"
    Flag =0
    Expression ="Schedule.SchID"
    Flag =0
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
        dbText "Name" ="Clients.ClientShNm"
        dbInteger "ColumnWidth" ="1590"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpBLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AtfRevq1.RevNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AtfRevq1.RevDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AtfRevq1.ReviewDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AtfRevq1.reviewer"
        dbInteger "ColumnWidth" ="1260"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.[client group]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientgroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BLCD"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =38
    Top =86
    Right =1544
    Bottom =793
    Left =-1
    Top =-1
    Right =1488
    Bottom =356
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="MstrLease"
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
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="AtfRevq1"
        Name =""
    End
    Begin
        Left =816
        Top =12
        Right =960
        Bottom =156
        Top =0
        Name ="SchGrp"
        Name =""
    End
End
