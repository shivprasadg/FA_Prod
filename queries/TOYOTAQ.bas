Operation =1
Option =0
Where ="(((Clients.ClientShNm) Like \"TOY*\"))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="Locations"
    Name ="ChildParts"
End
Begin OutputColumns
    Expression ="Clients.ClientShNm"
    Expression ="MstrLease.MLOrig"
    Expression ="MstrLease.MLNo"
    Expression ="Schedule.SchID"
    Expression ="Schedule.SchNo"
    Expression ="SchGrp.SGrpID"
    Expression ="SchGrp.SGrpNO"
    Expression ="Units.UnitUnitNum"
    Expression ="Units.UNITVIN"
    Expression ="Units.UnitID"
    Expression ="Units.GroupREF"
    Expression ="Units.UNITYR"
    Expression ="Units.UNITMAKE"
    Expression ="ChildParts.UnitREF"
    Expression ="ChildParts.ChildMake"
    Expression ="Locations.LocationName"
    Expression ="Locations.Locationsn"
    Expression ="Units.UNITFINALCOST"
    Expression ="ChildParts.childFinalCost"
    Expression ="Units.UnitEstDelDate"
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
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =2
    LeftTable ="Units"
    RightTable ="ChildParts"
    Expression ="Units.UnitID = ChildParts.UnitREF"
    Flag =2
End
Begin OrderBy
    Expression ="Schedule.SchID"
    Flag =0
    Expression ="SchGrp.SGrpNO"
    Flag =0
    Expression ="Units.UnitUnitNum"
    Flag =0
    Expression ="ChildParts.ChildMake"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbMemo "Filter" ="([TOYOTAQ].[Locationsn]=\"Long Beach\")"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbInteger "ColumnWidth" ="1905"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UNITYR"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildMake"
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
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.GroupREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UNITMAKE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UNITVIN"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.UnitREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.Locationsn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UNITFINALCOST"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.childFinalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitEstDelDate"
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
    Bottom =386
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
        Left =639
        Top =19
        Right =783
        Bottom =163
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =843
        Top =18
        Right =987
        Bottom =162
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =1057
        Top =254
        Right =1201
        Bottom =398
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =1094
        Top =63
        Right =1238
        Bottom =207
        Top =0
        Name ="ChildParts"
        Name =""
    End
End
