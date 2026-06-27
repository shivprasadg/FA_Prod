Operation =4
Option =0
Where ="(((Units.ClientGroupId) Is Null)) OR (((Units.clientId) Is Null))"
Begin InputTables
    Name ="Units"
    Name ="SchGrp"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="Clients"
End
Begin OutputColumns
    Name ="Units.ClientGroupId"
    Expression ="[Clients].[ClientGroupId]"
    Name ="Units.clientId"
    Expression ="[Clients].[clientID]"
End
Begin Joins
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
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
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
        dbText "Name" ="Units.ClientGroupId"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2265"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Clients.ClientGroupId"
        dbInteger "ColumnWidth" ="2430"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.clientId"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =97
    Top =69
    Right =1616
    Bottom =923
    Left =-1
    Top =-1
    Right =1487
    Bottom =556
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =233
        Top =120
        Right =377
        Bottom =530
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =32
        Top =499
        Right =176
        Bottom =643
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =27
        Top =208
        Right =171
        Bottom =344
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =28
        Top =347
        Right =172
        Bottom =491
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =23
        Top =59
        Right =167
        Bottom =203
        Top =0
        Name ="Clients"
        Name =""
    End
End
