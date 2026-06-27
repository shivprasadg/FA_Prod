Operation =1
Option =0
Where ="(((Clients.ClientShNm)=[enter short name]) AND ((Units.unitstatus)<>\"s\" And (U"
    "nits.unitstatus)<>\"r\"))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="Locations"
    Name ="Locations"
    Alias ="Locations_1"
    Name ="Locations"
    Alias ="Locations_2"
End
Begin OutputColumns
    Expression ="Units.unitvin"
    Expression ="Units.unitunitnum"
    Expression ="Clients.ClientShNm"
    Expression ="MstrLease.MLNo"
    Expression ="MstrLease.MLOrig"
    Expression ="SchGrp.SGrpNO"
    Expression ="SchGrp.SGrpNO"
    Expression ="SchGrp.Assignee_Bank"
    Expression ="Units.unityr"
    Expression ="Units.unitmake"
    Expression ="Units.unitmodel"
    Expression ="Units.unittype"
    Alias ="Domi cileLoc"
    Expression ="Locations.LocationName"
    Alias ="Carage City"
    Expression ="Locations.LCity"
    Alias ="DomS"
    Expression ="Locations.LState"
    Alias ="TitleSt"
    Expression ="Locations_1.LState"
    Alias ="RegSt"
    Expression ="Locations_2.LState"
    Expression ="Units.unitirp"
    Expression ="Units.unitstatus"
    Expression ="Units.UnitPlateNum"
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
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
    LeftTable ="Units"
    RightTable ="Locations"
    Expression ="Units.UnitGarageLocationRef = Locations.LocationID"
    Flag =2
    LeftTable ="Units"
    RightTable ="Locations_1"
    Expression ="Units.UnitTitleLocationRef = Locations_1.LocationID"
    Flag =2
    LeftTable ="Units"
    RightTable ="Locations_2"
    Expression ="Units.UnitRegLocationRef = Locations_2.LocationID"
    Flag =2
End
Begin OrderBy
    Expression ="Units.unitvin"
    Flag =0
    Expression ="Schedule.SchID"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbMemo "Filter" ="([Query1].[MLOrig]<>\"FA\" Or [Query1].[MLOrig] IS Null)"
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
        dbText "Name" ="MstrLease.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unityr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitmake"
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitmodel"
        dbInteger "ColumnWidth" ="2085"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unittype"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Carage City"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Domi cileLoc"
        dbInteger "ColumnWidth" ="3075"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitunitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitvin"
        dbInteger "ColumnWidth" ="2115"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DomS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TitleSt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RegSt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitirp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1005"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitstatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPlateNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.sgrpassignee"
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
    Bottom =298
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
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =816
        Top =12
        Right =960
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =978
        Top =13
        Right =1122
        Bottom =157
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =974
        Top =181
        Right =1118
        Bottom =325
        Top =0
        Name ="Locations_1"
        Name =""
    End
    Begin
        Left =757
        Top =180
        Right =901
        Bottom =324
        Top =0
        Name ="Locations_2"
        Name =""
    End
End
