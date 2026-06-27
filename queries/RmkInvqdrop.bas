Operation =1
Option =2
Where ="(((MstrLease.mlno) Not Like \"demo*\") AND ((Units.unitstatus)=\"R\" Or (Units.u"
    "nitstatus)=\"tr\" Or (Units.unitstatus)=\"k\" Or (Units.unitstatus)=\"s\" Or (Un"
    "its.unitstatus)=\"ks\") AND ((MstrLease.mlorig)=\"fa\" Or (MstrLease.mlorig)=\"r"
    "mkt\" Or (MstrLease.mlorig)=\"other\")) OR (((MstrLease.mlno) Not Like \"demo*\""
    ") AND ((Units.unitstatus)=\"R\" Or (Units.unitstatus)=\"tr\" Or (Units.unitstatu"
    "s)=\"k\" Or (Units.unitstatus)=\"S\" Or (Units.unitstatus)=\"KS\") AND ((MstrLea"
    "se.mlorig)=\"rmkt\") AND ((MstrLease.MLLessor)=\"fa llc\"))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="Locations"
    Name ="Rmkt"
    Name ="lookup_GroupMakes"
    Name ="Clients"
    Alias ="Clients_1"
End
Begin OutputColumns
    Expression ="Clients.clientshnm"
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
    LeftTable ="Units"
    RightTable ="Locations"
    Expression ="Units.UnitGarageLocationRef = Locations.LocationID"
    Flag =2
    LeftTable ="Units"
    RightTable ="Rmkt"
    Expression ="Units.UnitID = Rmkt.UnitRef"
    Flag =2
    LeftTable ="Units"
    RightTable ="lookup_GroupMakes"
    Expression ="Units.UnitMake = lookup_GroupMakes.groupMake"
    Flag =2
    LeftTable ="Rmkt"
    RightTable ="Clients_1"
    Expression ="Rmkt.BuyerRef = Clients_1.clientID"
    Flag =2
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
End
Begin OrderBy
    Expression ="Clients.clientshnm"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Clients.clientshnm"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1890"
        dbBoolean "ColumnHidden" ="0"
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
    Bottom =347
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
        Left =545
        Top =0
        Right =689
        Bottom =144
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =684
        Top =-19
        Right =835
        Bottom =153
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =910
        Top =166
        Right =1054
        Bottom =310
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =849
        Top =2
        Right =993
        Bottom =122
        Top =0
        Name ="Rmkt"
        Name =""
    End
    Begin
        Left =677
        Top =214
        Right =821
        Bottom =358
        Top =0
        Name ="lookup_GroupMakes"
        Name =""
    End
    Begin
        Left =1367
        Top =93
        Right =1511
        Bottom =237
        Top =0
        Name ="Clients_1"
        Name =""
    End
End
