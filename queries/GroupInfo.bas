Operation =1
Option =0
Begin InputTables
    Name ="SchGrp"
    Name ="vw_SixKeys"
    Name ="Lookup_LeaseStartsOn"
End
Begin OutputColumns
    Expression ="vw_SixKeys.clientID"
    Expression ="vw_SixKeys.MLOrig"
    Expression ="vw_SixKeys.MLID"
    Expression ="vw_SixKeys.SchID"
    Expression ="vw_SixKeys.MLNo"
    Expression ="vw_SixKeys.FAID"
    Expression ="SchGrp.SGrpID"
    Expression ="SchGrp.unitsingroup"
    Alias ="PLB"
    Expression ="IIf(IsNull([SgrpPlb]),No,[SgrpPlb])"
    Expression ="SchGrp.Assigned_Indicator"
    Expression ="SchGrp.Assignee_Bank"
    Expression ="SchGrp.AssignmentDate"
    Expression ="SchGrp.BLCD"
    Expression ="SchGrp.FundDate"
    Expression ="SchGrp.LeaseType"
    Expression ="SchGrp.LeaseTermPrimary"
    Expression ="SchGrp.BLCDSetBy"
    Expression ="SchGrp.SGrpNO"
    Expression ="SchGrp.GroupAddedDate"
    Expression ="SchGrp.PerDiemStartsOn"
    Alias ="PDStart"
    Expression ="Lookup_LeaseStartsOn.StartsOn"
    Expression ="SchGrp.SgrpCommInd"
    Expression ="SchGrp.OrigEquipCost"
    Expression ="SchGrp.LeasePymtForGroup"
    Expression ="SchGrp.ExchangeITGroup"
    Expression ="SchGrp.SgrpDesc"
End
Begin Joins
    LeftTable ="SchGrp"
    RightTable ="vw_SixKeys"
    Expression ="SchGrp.SGrpID = vw_SixKeys.SGrpID"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Lookup_LeaseStartsOn"
    Expression ="SchGrp.PerDiemStartsOn = Lookup_LeaseStartsOn.ID"
    Flag =2
End
Begin OrderBy
    Expression ="vw_SixKeys.SchID"
    Flag =1
    Expression ="vw_SixKeys.MLNo"
    Flag =0
End
Begin Groups
    Expression ="vw_SixKeys.clientID"
    GroupLevel =0
    Expression ="vw_SixKeys.MLOrig"
    GroupLevel =0
    Expression ="vw_SixKeys.MLID"
    GroupLevel =0
    Expression ="vw_SixKeys.SchID"
    GroupLevel =0
    Expression ="vw_SixKeys.MLNo"
    GroupLevel =0
    Expression ="vw_SixKeys.FAID"
    GroupLevel =0
    Expression ="SchGrp.SGrpID"
    GroupLevel =0
    Expression ="SchGrp.unitsingroup"
    GroupLevel =0
    Expression ="IIf(IsNull([SgrpPlb]),No,[SgrpPlb])"
    GroupLevel =0
    Expression ="SchGrp.Assigned_Indicator"
    GroupLevel =0
    Expression ="SchGrp.Assignee_Bank"
    GroupLevel =0
    Expression ="SchGrp.AssignmentDate"
    GroupLevel =0
    Expression ="SchGrp.BLCD"
    GroupLevel =0
    Expression ="SchGrp.FundDate"
    GroupLevel =0
    Expression ="SchGrp.LeaseType"
    GroupLevel =0
    Expression ="SchGrp.LeaseTermPrimary"
    GroupLevel =0
    Expression ="SchGrp.BLCDSetBy"
    GroupLevel =0
    Expression ="SchGrp.SGrpNO"
    GroupLevel =0
    Expression ="SchGrp.GroupAddedDate"
    GroupLevel =0
    Expression ="SchGrp.PerDiemStartsOn"
    GroupLevel =0
    Expression ="Lookup_LeaseStartsOn.StartsOn"
    GroupLevel =0
    Expression ="SchGrp.SgrpCommInd"
    GroupLevel =0
    Expression ="SchGrp.OrigEquipCost"
    GroupLevel =0
    Expression ="SchGrp.LeasePymtForGroup"
    GroupLevel =0
    Expression ="SchGrp.ExchangeITGroup"
    GroupLevel =0
    Expression ="SchGrp.SgrpDesc"
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
        dbText "Name" ="PLB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.Assigned_Indicator"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.unitsingroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.FundDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.Assignee_Bank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.AssignmentDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LeaseType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LeaseTermPrimary"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BLCDSetBy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.OrigEquipCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.GroupAddedDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpCommInd"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1725"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="SchGrp.PerDiemStartsOn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LeasePymtForGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.ExchangeITGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.[Assigned_Indicator]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDStart"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lookup_LeaseStartsOn.StartsOn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lookup_LeaseStartsOn_1.StartsOn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLCDStart"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =252
    Top =692
    Right =1839
    Bottom =1525
    Left =-1
    Top =-1
    Right =1555
    Bottom =345
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =214
        Top =22
        Right =465
        Bottom =354
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =34
        Top =11
        Right =178
        Bottom =374
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =606
        Top =67
        Right =750
        Bottom =211
        Top =0
        Name ="Lookup_LeaseStartsOn"
        Name =""
    End
End
