Operation =1
Option =0
Having ="(((UnitExitOption.OptionTypeID)=1))"
Begin InputTables
    Name ="UnitExitOption"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Expression ="UnitExitOption.UnitID"
    Expression ="vw_SixKeys.SchID"
    Expression ="vw_SixKeys.SGrpID"
    Expression ="vw_SixKeys.Schedule"
    Expression ="vw_SixKeys.UnitGroup"
    Expression ="UnitExitOption.OptionTypeID"
    Expression ="UnitExitOption.ReplacedByScheduleID"
    Expression ="UnitExitOption.ReplacedByGroupID"
    Expression ="UnitExitOption.ReplacedByUnitID"
    Alias ="NewSchedule"
    Expression ="Schedule.SchNo"
    Alias ="NewGroup"
    Expression ="SchGrp.SGrpNO"
    Alias ="NewAssetID"
    Expression ="Units.UnitID"
    Alias ="NewUnitNum"
    Expression ="Units.UnitUnitNum"
End
Begin Joins
    LeftTable ="UnitExitOption"
    RightTable ="Schedule"
    Expression ="UnitExitOption.ReplacedByScheduleID = Schedule.SchID"
    Flag =1
    LeftTable ="UnitExitOption"
    RightTable ="SchGrp"
    Expression ="UnitExitOption.ReplacedByGroupID = SchGrp.SGrpID"
    Flag =1
    LeftTable ="UnitExitOption"
    RightTable ="Units"
    Expression ="UnitExitOption.ReplacedByUnitID = Units.UnitID"
    Flag =2
    LeftTable ="UnitExitOption"
    RightTable ="vw_SixKeys"
    Expression ="UnitExitOption.UnitID = vw_SixKeys.AssetID"
    Flag =1
End
Begin Groups
    Expression ="UnitExitOption.UnitID"
    GroupLevel =0
    Expression ="vw_SixKeys.SchID"
    GroupLevel =0
    Expression ="vw_SixKeys.SGrpID"
    GroupLevel =0
    Expression ="vw_SixKeys.Schedule"
    GroupLevel =0
    Expression ="vw_SixKeys.UnitGroup"
    GroupLevel =0
    Expression ="UnitExitOption.OptionTypeID"
    GroupLevel =0
    Expression ="UnitExitOption.ReplacedByScheduleID"
    GroupLevel =0
    Expression ="UnitExitOption.ReplacedByGroupID"
    GroupLevel =0
    Expression ="UnitExitOption.ReplacedByUnitID"
    GroupLevel =0
    Expression ="Schedule.SchNo"
    GroupLevel =0
    Expression ="SchGrp.SGrpNO"
    GroupLevel =0
    Expression ="Units.UnitID"
    GroupLevel =0
    Expression ="Units.UnitUnitNum"
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
        dbText "Name" ="UnitExitOption.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.OptionTypeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.ReplacedByGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.ReplacedByScheduleID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.ReplacedByUnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewAssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewSchedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.SGrpID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =264
    Top =87
    Right =1592
    Bottom =968
    Left =-1
    Top =-1
    Right =1296
    Bottom =478
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =73
        Top =80
        Right =388
        Bottom =497
        Top =0
        Name ="UnitExitOption"
        Name =""
    End
    Begin
        Left =448
        Top =83
        Right =592
        Bottom =227
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =473
        Top =296
        Right =617
        Bottom =440
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =677
        Top =425
        Right =821
        Bottom =569
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =665
        Top =6
        Right =809
        Bottom =351
        Top =0
        Name ="vw_Client_FAID_PO_Lease_Sch_Grp_Units"
        Name =""
    End
End
