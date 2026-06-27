Operation =1
Option =2
Where ="(((MstrLease.MLOrig)=\"fa\") AND ((SchGrp.SGrpRALb)>=[enter start dt] And (SchGr"
    "p.SGrpRALb)<=[enterend date])) OR (((SchGrp.SGrpRALO)>=[enter start dt] And (Sch"
    "Grp.SGrpRALO)<=[enterend date])) OR (((Clients.clientCompanyName)=\"Southern Win"
    "e & Spirits of America, Inc.\")) OR (((Clients.clientCompanyName) Like \"rein*\""
    "))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
End
Begin OutputColumns
    Expression ="Clients.clientCompanyName"
    Expression ="MstrLease.MLOrig"
    Expression ="Schedule.schno"
    Expression ="SchGrp.SGrpRALO"
    Expression ="SchGrp.SGrpRALb"
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
        dbText "Name" ="Clients.clientCompanyName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3885"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="MstrLease.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpRALO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpRALb"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.schno"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =114
    Top =178
    Right =1248
    Bottom =755
    Left =-1
    Top =-1
    Right =1102
    Bottom =226
    Left =1103
    Top =0
    ColumnsShown =539
    Begin
        Left =-1079
        Top =-12
        Right =-936
        Bottom =134
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =-924
        Top =-5
        Right =-780
        Bottom =139
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =-734
        Top =11
        Right =-590
        Bottom =155
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =-542
        Top =12
        Right =-398
        Bottom =156
        Top =0
        Name ="SchGrp"
        Name =""
    End
End
