Operation =1
Option =0
Having ="(((Schedule.SchFAID)<>\"\"))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
End
Begin OutputColumns
    Expression ="Clients.ClientGroupId"
    Expression ="Clients.clientID"
    Expression ="Clients.ClientShNm"
    Alias ="FAID"
    Expression ="Schedule.SchFAID"
    Expression ="MstrLease.MLID"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="MstrLease"
    Expression ="Clients.clientID = MstrLease.Client"
    Flag =1
    LeftTable ="MstrLease"
    RightTable ="Schedule"
    Expression ="MstrLease.MLID = Schedule.SchMLRef"
    Flag =1
End
Begin OrderBy
    Expression ="Clients.ClientShNm"
    Flag =0
End
Begin Groups
    Expression ="Clients.ClientGroupId"
    GroupLevel =0
    Expression ="Clients.clientID"
    GroupLevel =0
    Expression ="Clients.ClientShNm"
    GroupLevel =0
    Expression ="Schedule.SchFAID"
    GroupLevel =0
    Expression ="MstrLease.MLID"
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
        dbText "Name" ="Clients.clientID"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =541
    Top =41
    Right =1884
    Bottom =619
    Left =-1
    Top =-1
    Right =1311
    Bottom =196
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =37
        Top =36
        Right =181
        Bottom =180
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =217
        Top =30
        Right =361
        Bottom =174
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =388
        Top =52
        Right =532
        Bottom =196
        Top =0
        Name ="Schedule"
        Name =""
    End
End
