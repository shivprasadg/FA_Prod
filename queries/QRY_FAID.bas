Operation =1
Option =0
Begin InputTables
    Name ="SchGrp"
    Name ="Schedule"
    Name ="Clients"
    Name ="MstrLease"
    Name ="Units"
    Name ="Rmkt"
End
Begin OutputColumns
    Expression ="Clients.ClientShNm"
    Expression ="Units.UnitID"
    Expression ="Units.UnitUnitNum"
    Expression ="Schedule.SchFAID"
End
Begin Joins
    LeftTable ="SchGrp"
    RightTable ="Schedule"
    Expression ="SchGrp.SGrpID = Schedule.SchID"
    Flag =3
    LeftTable ="Clients"
    RightTable ="MstrLease"
    Expression ="Clients.clientID = MstrLease.Client"
    Flag =2
    LeftTable ="Schedule"
    RightTable ="MstrLease"
    Expression ="Schedule.SchMLRef = MstrLease.MLID"
    Flag =3
    LeftTable ="Units"
    RightTable ="Rmkt"
    Expression ="Units.UnitID = Rmkt.UnitRef"
    Flag =2
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbMemo "OrderBy" ="[Query4].[ClientShNm]"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbInteger "ColumnWidth" ="2520"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchFAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =38
    Top =86
    Right =1214
    Bottom =872
    Left =-1
    Top =-1
    Right =1144
    Bottom =452
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =612
        Top =58
        Right =756
        Bottom =432
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =400
        Top =59
        Right =544
        Bottom =455
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =0
        Top =5
        Right =144
        Bottom =149
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =204
        Top =21
        Right =348
        Bottom =165
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =828
        Top =35
        Right =999
        Bottom =450
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =1053
        Top =35
        Right =1197
        Bottom =280
        Top =0
        Name ="Rmkt"
        Name =""
    End
End
