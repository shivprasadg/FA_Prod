Operation =1
Option =0
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="MakeModelManager_WorkTable"
End
Begin OutputColumns
    Expression ="Clients.ClientShNm"
    Expression ="MstrLease.MLNo"
    Alias ="SchGrp"
    Expression ="[SchNo] & ' |  ' & [SGrpNO]"
    Expression ="Units.UnitID"
    Expression ="Units.UnitUnitNum"
    Expression ="Units.UnitStatus"
    Expression ="Units.UnitYr"
    Expression ="Units.UnitSubType"
    Expression ="Units.UnitType"
    Expression ="Units.UnitMake"
    Expression ="Units.UnitModel"
    Expression ="Units.UnitUOM"
    Expression ="Units.UnitVin"
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
    RightTable ="MakeModelManager_WorkTable"
    Expression ="Units.UnitID = MakeModelManager_WorkTable.UnitID"
    Flag =1
End
Begin OrderBy
    Expression ="Clients.ClientShNm"
    Flag =0
    Expression ="[SchNo] & ' |  ' & [SGrpNO]"
    Flag =0
    Expression ="Units.UnitID"
    Flag =0
    Expression ="Units.UnitType"
    Flag =0
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
        dbText "Name" ="SchGrp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitSubType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUOM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVin"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =32
    Top =87
    Right =1425
    Bottom =1085
    Left =-1
    Top =-1
    Right =1369
    Bottom =459
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =880
        Top =67
        Right =1024
        Bottom =211
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =705
        Top =49
        Right =849
        Bottom =193
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =533
        Top =31
        Right =677
        Bottom =175
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =362
        Top =13
        Right =506
        Bottom =157
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =188
        Top =13
        Right =332
        Bottom =468
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =16
        Top =13
        Right =160
        Bottom =90
        Top =0
        Name ="MakeModelManager_WorkTable"
        Name =""
    End
End
