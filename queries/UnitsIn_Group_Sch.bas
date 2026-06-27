Operation =1
Option =0
Having ="(((Units.GroupREF)=[Forms]![GroupForm]![GroupID]))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="Locations"
    Name ="lookup_GroupMakes"
    Name ="ChildParts"
    Name ="lookup_groupTypes"
    Name ="LookUp_UnitDesc"
End
Begin OutputColumns
    Expression ="Units.UnitID"
    Expression ="Clients.ClientShNm"
    Expression ="MstrLease.MLNo"
    Expression ="Schedule.SchNo"
    Expression ="SchGrp.SGrpNO"
    Expression ="Units.GroupREF"
    Expression ="SchGrp.SGrpID"
    Expression ="Units.UnitStatus"
    Expression ="Units.UnitEstDelDate"
    Expression ="Units.UnitVIN"
    Expression ="Units.UnitUnitNum"
    Expression ="Units.UnitType"
    Expression ="Units.UnitSubType"
    Alias ="vin4"
    Expression ="(Right([UNITVIN],6))"
    Expression ="lookup_groupTypes.GrpTypAbrev"
    Expression ="Units.UnitYr"
    Expression ="Units.UnitMake"
    Expression ="Units.UnitModel"
    Expression ="Units.UnitCOASent"
    Expression ="Units.UnitCOABack"
    Expression ="Units.UnitAcceptDate"
    Expression ="Units.UnitFinalCost"
    Alias ="SumOfchildFinalCost"
    Expression ="Sum(ChildParts.childFinalCost)"
    Expression ="Units.unitoffleasedt"
    Expression ="Units.UnitActualDelDate"
    Alias ="ULoc"
    Expression ="\"(\" & [lOCATIONSN] & \")\" & \" \" & [LCity] & \" \" & [locations].[LState]"
    Expression ="lookup_GroupMakes.MakeAbrev"
    Expression ="Units.UnitDocAccpt"
    Expression ="Units.unitinserv"
    Expression ="Units.UnitInServInd"
    Expression ="Units.UnitPlateNum"
    Expression ="Units.UnitOLInd"
    Expression ="Units.unitunitnum2"
    Expression ="Units.Axle"
    Expression ="Schedule.SchID"
    Expression ="SchGrp.SGrpID"
    Alias ="UnitDesc"
    Expression ="LookUp_UnitDesc.DescriptionLong"
    Alias ="CofABck-Cpy"
    Expression ="Nz([units].[UnitCoABack],[Units].[UnitCofACopy])"
    Expression ="Units.UnitCOABack"
    Expression ="Units.UnitCofACopy"
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
    RightTable ="lookup_GroupMakes"
    Expression ="Units.UnitMake = lookup_GroupMakes.groupMake"
    Flag =2
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
    LeftTable ="Units"
    RightTable ="ChildParts"
    Expression ="Units.UnitID = ChildParts.UnitREF"
    Flag =2
    LeftTable ="Units"
    RightTable ="lookup_groupTypes"
    Expression ="Units.UnitType = lookup_groupTypes.groupType"
    Flag =2
    LeftTable ="Units"
    RightTable ="LookUp_UnitDesc"
    Expression ="Units.UnitID = LookUp_UnitDesc.AssetID"
    Flag =1
End
Begin OrderBy
    Expression ="Units.UnitID"
    Flag =0
    Expression ="Units.UnitType"
    Flag =0
End
Begin Groups
    Expression ="Units.UnitID"
    GroupLevel =0
    Expression ="Clients.ClientShNm"
    GroupLevel =0
    Expression ="MstrLease.MLNo"
    GroupLevel =0
    Expression ="Schedule.SchNo"
    GroupLevel =0
    Expression ="SchGrp.SGrpNO"
    GroupLevel =0
    Expression ="Units.GroupREF"
    GroupLevel =0
    Expression ="SchGrp.SGrpID"
    GroupLevel =0
    Expression ="Units.UnitStatus"
    GroupLevel =0
    Expression ="Units.UnitEstDelDate"
    GroupLevel =0
    Expression ="Units.UnitVIN"
    GroupLevel =0
    Expression ="Units.UnitUnitNum"
    GroupLevel =0
    Expression ="Units.UnitType"
    GroupLevel =0
    Expression ="Units.UnitSubType"
    GroupLevel =0
    Expression ="(Right([UNITVIN],6))"
    GroupLevel =0
    Expression ="lookup_groupTypes.GrpTypAbrev"
    GroupLevel =0
    Expression ="Units.UnitYr"
    GroupLevel =0
    Expression ="Units.UnitMake"
    GroupLevel =0
    Expression ="Units.UnitModel"
    GroupLevel =0
    Expression ="Units.UnitCOASent"
    GroupLevel =0
    Expression ="Units.UnitAcceptDate"
    GroupLevel =0
    Expression ="Units.UnitFinalCost"
    GroupLevel =0
    Expression ="Units.unitoffleasedt"
    GroupLevel =0
    Expression ="Units.UnitActualDelDate"
    GroupLevel =0
    Expression ="\"(\" & [lOCATIONSN] & \")\" & \" \" & [LCity] & \" \" & [locations].[LState]"
    GroupLevel =0
    Expression ="lookup_GroupMakes.MakeAbrev"
    GroupLevel =0
    Expression ="Units.UnitDocAccpt"
    GroupLevel =0
    Expression ="Units.unitinserv"
    GroupLevel =0
    Expression ="Units.UnitInServInd"
    GroupLevel =0
    Expression ="Units.UnitPlateNum"
    GroupLevel =0
    Expression ="Units.UnitOLInd"
    GroupLevel =0
    Expression ="Units.unitunitnum2"
    GroupLevel =0
    Expression ="Units.Axle"
    GroupLevel =0
    Expression ="Schedule.SchID"
    GroupLevel =0
    Expression ="SchGrp.SGrpID"
    GroupLevel =0
    Expression ="LookUp_UnitDesc.DescriptionLong"
    GroupLevel =0
    Expression ="Units.UnitCOABack"
    GroupLevel =0
    Expression ="Units.UnitCofACopy"
    GroupLevel =0
    Expression ="SchGrp.SGrpID"
    GroupLevel =0
    Expression ="Units.UnitCOABack"
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
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.GroupREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitEstDelDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vin4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lookup_groupTypes.GrpTypAbrev"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitCOASent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitCOABack"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitAcceptDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitFinalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfchildFinalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitoffleasedt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitActualDelDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ULoc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lookup_GroupMakes.MakeAbrev"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDocAccpt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitinserv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInServInd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPlateNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitOLInd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitunitnum2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.Axle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitSubType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitCofACopy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofABck-Cpy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LookUp_UnitDesc.DescriptionLong"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1006"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1019"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-31
    Top =252
    Right =1458
    Bottom =1242
    Left =-1
    Top =-1
    Right =1471
    Bottom =311
    Left =0
    Top =0
    ColumnsShown =543
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
        Left =1008
        Top =12
        Right =1152
        Bottom =156
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =48
        Top =156
        Right =192
        Bottom =300
        Top =0
        Name ="lookup_GroupMakes"
        Name =""
    End
    Begin
        Left =240
        Top =156
        Right =384
        Bottom =300
        Top =0
        Name ="ChildParts"
        Name =""
    End
    Begin
        Left =432
        Top =156
        Right =576
        Bottom =300
        Top =0
        Name ="lookup_groupTypes"
        Name =""
    End
    Begin
        Left =1181
        Top =12
        Right =1344
        Bottom =198
        Top =0
        Name ="LookUp_UnitDesc"
        Name =""
    End
End
