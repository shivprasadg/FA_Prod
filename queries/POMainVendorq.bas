Operation =1
Option =0
Where ="(((Childinv.childvendorref)=[vendorref]) AND ((MstrLease.MLOrig)=\"fa\"))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="Childinv"
End
Begin OutputColumns
    Expression ="Clients.ClientGroupID"
    Expression ="Clients.ClientShNm"
    Expression ="MstrLease.MLNo"
    Expression ="Schedule.SchNo"
    Alias ="FAID"
    Expression ="IIf(IsNull([unitfaid]),IIf(IsNull([sgrpfaid]),[schfaid],[sgrpfaid]),[unitfaid])"
    Expression ="Units.VendorREF"
    Expression ="Units.unityr"
    Expression ="Units.unitmake"
    Expression ="Units.unitmodel"
    Expression ="Units.unittype"
    Expression ="Units.Axle"
    Expression ="Units.unitfinalcost"
    Expression ="Childinv.ChildYear"
    Expression ="Childinv.ChildMake"
    Expression ="Childinv.ChildClass"
    Expression ="Childinv.ChildType"
    Expression ="Childinv.childfinalcost"
    Expression ="Childinv.childvendorref"
End
Begin Joins
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
    RightTable ="Childinv"
    Expression ="Units.UnitID = Childinv.UnitID"
    Flag =1
    LeftTable ="Clients"
    RightTable ="MstrLease"
    Expression ="Clients.clientID = MstrLease.Client"
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
        dbText "Name" ="Units.unityr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unittype"
        dbInteger "ColumnWidth" ="2640"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.Axle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitmake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitmodel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Childinv.ChildYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Childinv.ChildMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Childinv.ChildModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Childinv.ChildType"
        dbInteger "ColumnWidth" ="2565"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Childinv.childfinalcost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitfinalcost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.VendorREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Childinv.childvendorref"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
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
    Bottom =506
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =96
        Top =192
        Right =240
        Bottom =336
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =389
        Top =34
        Right =533
        Bottom =178
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =243
        Top =136
        Right =387
        Bottom =280
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =436
        Top =199
        Right =580
        Bottom =343
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =647
        Top =188
        Right =791
        Bottom =332
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =860
        Top =200
        Right =1004
        Bottom =344
        Top =0
        Name ="Childinv"
        Name =""
    End
End
