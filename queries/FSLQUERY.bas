Operation =1
Option =0
Where ="(((Clients.ClientShNm) Like \"costco*\"))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="FSL"
End
Begin OutputColumns
    Expression ="Clients.ClientShNm"
    Expression ="Schedule.SchNo"
    Expression ="Schedule.SchFSL"
    Expression ="FSL.ID"
    Expression ="Schedule.SchShortDesc"
    Expression ="FSL.FSLYR1"
    Expression ="FSL.FSLYR2"
    Expression ="FSL.FSLYR3"
    Expression ="FSL.FSLYR4"
    Expression ="FSL.FSLYR5"
    Expression ="FSL.FSLYR6"
    Expression ="FSL.FSLYR7"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="MstrLease"
    Expression ="Clients.clientID = MstrLease.Client"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="FSL"
    Expression ="Schedule.FSLSCH = FSL.ID"
    Flag =2
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
        dbText "Name" ="FSL.FSLYR1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FSL.FSLYR2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FSL.FSLYR3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FSL.FSLYR4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FSL.FSLYR5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FSL.FSLYR6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FSL.FSLYR7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FSL.ID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1140"
        dbInteger "ColumnOrder" ="5"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Schedule.SchShortDesc"
        dbInteger "ColumnWidth" ="4485"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbInteger "ColumnWidth" ="2400"
        dbInteger "ColumnOrder" ="2"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchFSL"
        dbInteger "ColumnWidth" ="1035"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbInteger "ColumnWidth" ="1530"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1258
    Bottom =686
    Left =-1
    Top =-1
    Right =1226
    Bottom =339
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
        Left =664
        Top =39
        Right =808
        Bottom =236
        Top =0
        Name ="FSL"
        Name =""
    End
End
