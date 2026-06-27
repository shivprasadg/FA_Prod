Operation =1
Option =0
Begin InputTables
    Name ="SchGrp"
    Name ="UnitTotalOEC"
End
Begin OutputColumns
    Alias ="ScheduleID"
    Expression ="SchGrp.SGrpSchID"
    Expression ="UnitTotalOEC.GroupID"
    Expression ="UnitTotalOEC.AssetID"
    Alias ="StripsNum"
    Expression ="IIf(IsDate([BLCD]) And IsDate([AssignmentDate]),DateDiff(\"m\",[BLCD],[Assignmen"
        "tDate]),0)"
    Expression ="UnitTotalOEC.SalesRent"
    Alias ="RebateUnitActual"
    Expression ="SchGrp.SgrpSynRebate"
    Expression ="SchGrp.RebateUnitProjected"
    Alias ="UTA"
    Expression ="CCur(Nz([SgrpSynRebate],0))"
End
Begin Joins
    LeftTable ="SchGrp"
    RightTable ="UnitTotalOEC"
    Expression ="SchGrp.SGrpID = UnitTotalOEC.GroupID"
    Flag =1
End
Begin Groups
    Expression ="SchGrp.SGrpSchID"
    GroupLevel =0
    Expression ="UnitTotalOEC.GroupID"
    GroupLevel =0
    Expression ="UnitTotalOEC.AssetID"
    GroupLevel =0
    Expression ="IIf(IsDate([BLCD]) And IsDate([AssignmentDate]),DateDiff(\"m\",[BLCD],[Assignmen"
        "tDate]),0)"
    GroupLevel =0
    Expression ="UnitTotalOEC.SalesRent"
    GroupLevel =0
    Expression ="SchGrp.SgrpSynRebate"
    GroupLevel =0
    Expression ="SchGrp.RebateUnitProjected"
    GroupLevel =0
    Expression ="CCur(Nz([SgrpSynRebate],0))"
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
        dbText "Name" ="UnitTotalOEC.GroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTotalOEC.SalesRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StripsNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RebateUnitActual"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.RebateUnitProjected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UTA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTotalOEC.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =2
    Left =-8
    Top =-31
    Right =1293
    Bottom =1034
    Left =-1
    Top =-1
    Right =1269
    Bottom =527
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =153
        Top =62
        Right =297
        Bottom =206
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =362
        Top =216
        Right =506
        Bottom =458
        Top =0
        Name ="UnitTotalOEC"
        Name =""
    End
End
