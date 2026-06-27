Operation =1
Option =0
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
End
Begin OutputColumns
    Expression ="Clients.ClientShNm"
    Expression ="MstrLease.MLNo"
    Expression ="Schedule.SchNo"
    Expression ="SchGrp.*"
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
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
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
        dbText "Name" ="Clients.ClientShNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =21
    Top =80
    Right =1008
    Bottom =604
    Left =-1
    Top =-1
    Right =955
    Bottom =221
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =60
        Top =-2
        Right =204
        Bottom =142
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =234
        Top =3
        Right =378
        Bottom =147
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =392
        Top =1
        Right =536
        Bottom =145
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =556
        Top =5
        Right =700
        Bottom =146
        Top =0
        Name ="SchGrp"
        Name =""
    End
End
