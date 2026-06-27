Operation =1
Option =0
Where ="(((MstrLease.MLOrig)=\"fa\"))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
End
Begin OutputColumns
    Expression ="Clients.clientID"
    Expression ="Clients.ClientShNm"
    Expression ="Clients.clientCompanyName"
    Expression ="MstrLease.MLOrig"
    Expression ="Schedule.SchNo"
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
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientCompanyName"
        dbInteger "ColumnWidth" ="2625"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbInteger "ColumnWidth" ="2505"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1698
    Bottom =855
    Left =-1
    Top =-1
    Right =1666
    Bottom =481
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =15
        Top =11
        Right =230
        Bottom =155
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =254
        Top =11
        Right =398
        Bottom =155
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =424
        Top =10
        Right =568
        Bottom =154
        Top =0
        Name ="Schedule"
        Name =""
    End
End
