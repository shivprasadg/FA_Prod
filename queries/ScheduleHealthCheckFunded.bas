Operation =1
Option =0
Where ="((Not (vw_ScheduleHealthCheck.FundX) Is Null And (vw_ScheduleHealthCheck.FundX)>"
    "#12/31/2021#))"
Begin InputTables
    Name ="vw_ScheduleHealthCheck"
End
Begin OutputColumns
    Expression ="vw_ScheduleHealthCheck.*"
End
Begin OrderBy
    Expression ="vw_ScheduleHealthCheck.Lessee"
    Flag =0
    Expression ="vw_ScheduleHealthCheck.Schedule"
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
        dbText "Name" ="vw_ScheduleHealthCheck.FundX"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleHealthCheck.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleHealthCheck.Lessee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleHealthCheck.SchOutBack"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleHealthCheck.TMTSigned"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleHealthCheck.COA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleHealthCheck.Delvery"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleHealthCheck.FinalEconRecv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleHealthCheck.OPSPkg"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleHealthCheck.RALOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleHealthCheck.RALBack"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleHealthCheck.PermPkgOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleHealthCheck.FATitleTrust"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleHealthCheck.IntermFundBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleHealthCheck.IntermFund"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleHealthCheck.BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleHealthCheck.Assigned"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleHealthCheck.Assignee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleHealthCheck.FundDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleHealthCheck.Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleHealthCheck.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleHealthCheck.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ScheduleHealthCheck.SchID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =94
    Top =4
    Right =1585
    Bottom =1226
    Left =-1
    Top =-1
    Right =1467
    Bottom =842
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =62
        Right =427
        Bottom =394
        Top =0
        Name ="vw_ScheduleHealthCheck"
        Name =""
    End
End
