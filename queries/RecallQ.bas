Operation =1
Option =0
Where ="(((Recalls.RCSent) Is Null)) OR (((Recalls.RCConf) Is Null)) OR (((Recalls.RCIns"
    "pDt)=\"isnull\") AND ((Recalls.rcrepairreq) Is Null)) OR (((Recalls.rcrepairreq)"
    "=Yes) AND ((Recalls.rcrepdt) Is Null))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="Recalls"
    Name ="Locations"
End
Begin OutputColumns
    Expression ="Recalls.RCID"
    Alias ="Grp"
    Expression ="[rcdt] & [rcno] & [rcdesc]"
    Expression ="Recalls.RCDt"
    Expression ="Recalls.RCNo"
    Expression ="Recalls.RCDesc"
    Expression ="Clients.ClientShNm"
    Expression ="Locations.LocationName"
    Expression ="MstrLease.MLNo"
    Expression ="Schedule.SchNo"
    Expression ="SchGrp.SGrpNO"
    Expression ="SchGrp.Assignee_Bank"
    Expression ="Units.UnitUnitNum"
    Expression ="Units.unitvin"
    Expression ="Recalls.RCMake"
    Expression ="Recalls.RCCompMake"
    Expression ="Recalls.RCRec"
    Expression ="Recalls.RCSent"
    Expression ="Recalls.RCConf"
    Alias ="ROrder"
    Expression ="IIf(IsNull([rcsent]),1,IIf(IsNull([rcconf]),2,IIf(IsNull([rcinspdt]),3,IIf(IsNul"
        "l([rcrepairreq]),4,IIf(IsNull([rcrepdt]),5,\"ERR\")))))"
    Expression ="Recalls.RCInspDt"
    Expression ="Recalls.rcrepairreq"
    Expression ="Recalls.rcrepdt"
    Expression ="Locations.Locationsn"
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
    RightTable ="Recalls"
    Expression ="Units.UnitID = Recalls.RCUnitID"
    Flag =1
    LeftTable ="Units"
    RightTable ="Locations"
    Expression ="Units.UnitGarageLocationRef = Locations.LocationID"
    Flag =1
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
        dbText "Name" ="Recalls.RCID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Recalls.RCDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Recalls.RCNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Recalls.RCDesc"
        dbInteger "ColumnWidth" ="12375"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationName"
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
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpAssignee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitvin"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Recalls.RCMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Recalls.RCCompMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Recalls.RCRec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Recalls.RCSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Recalls.RCConf"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ROrder"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Recalls.RCInspDt"
        dbInteger "ColumnWidth" ="9120"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Recalls.rcrepairreq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Recalls.rcrepdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.Locationsn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Grp"
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
    Bottom =252
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
        Left =48
        Top =156
        Right =192
        Bottom =300
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =240
        Top =156
        Right =384
        Bottom =300
        Top =0
        Name ="Recalls"
        Name =""
    End
    Begin
        Left =432
        Top =156
        Right =576
        Bottom =300
        Top =0
        Name ="Locations"
        Name =""
    End
End
