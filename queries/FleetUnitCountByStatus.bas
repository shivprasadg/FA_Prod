Operation =1
Option =0
Begin InputTables
    Name ="Units"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="MstrLease"
    Name ="Clients"
End
Begin OutputColumns
    Expression ="Clients.clientID"
    Alias ="Status"
    Expression ="Units.UnitStatus"
    Alias ="UnitCount"
    Expression ="Count(Units.UnitID)"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="MstrLease"
    Expression ="Clients.clientID = MstrLease.Client"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
    LeftTable ="MstrLease"
    RightTable ="Schedule"
    Expression ="MstrLease.MLID = Schedule.SchMLRef"
    Flag =1
End
Begin Groups
    Expression ="Clients.clientID"
    GroupLevel =0
    Expression ="Units.UnitStatus"
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
dbMemo "Filter" ="([FleetUnitCountByStatus].[clientID]=3632)"
Begin
    Begin
        dbText "Name" ="UnitCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =199
    Top =132
    Right =1673
    Bottom =1064
    Left =-1
    Top =-1
    Right =1442
    Bottom =544
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =728
        Top =180
        Right =872
        Bottom =382
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =402
        Top =177
        Right =546
        Bottom =366
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =554
        Top =76
        Right =692
        Bottom =214
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =235
        Top =10
        Right =379
        Bottom =154
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =57
        Top =23
        Right =201
        Bottom =167
        Top =0
        Name ="Clients"
        Name =""
    End
End
