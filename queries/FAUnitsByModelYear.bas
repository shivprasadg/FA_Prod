Operation =1
Option =0
Where ="(((vw_SixKeys.MLOrig)=\"FA\"))"
Begin InputTables
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Alias ="ModelYear"
    Expression ="vw_SixKeys.MYear"
    Expression ="vw_SixKeys.Make"
    Expression ="vw_SixKeys.Type"
    Alias ="Units"
    Expression ="Count(vw_SixKeys.UnitID)"
End
Begin OrderBy
    Expression ="vw_SixKeys.MYear"
    Flag =1
    Expression ="vw_SixKeys.Make"
    Flag =0
    Expression ="vw_SixKeys.Type"
    Flag =0
End
Begin Groups
    Expression ="vw_SixKeys.MYear"
    GroupLevel =0
    Expression ="vw_SixKeys.Make"
    GroupLevel =0
    Expression ="vw_SixKeys.Type"
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
dbBoolean "TotalsRow" ="-1"
Begin
    Begin
        dbText "Name" ="LeaseReport_Syndicated_V_OnOrder.OrderedYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LeaseReport_Syndicated_V_OnOrder.OnOrderOEC"
        dbLong "AggregateType" ="0"
    End
    Begin
        dbText "Name" ="OrderedOwnedSyn_LeaseReport.Owned"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SydCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DelveryYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LeaseReport_Syndicated_V_OnOrder.SYDYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LeaseReport_Syndicated_V_OnOrder.SydOEC"
        dbLong "AggregateType" ="0"
    End
    Begin
        dbText "Name" ="OrderedOwnedSyn_LeaseReport.Make"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderedOwnedSyn_LeaseReport.OwnedCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderedOwnedSyn_LeaseReport.SYDYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderedOwnedSyn_LeaseReport_1.OrderedYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.MYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LeaseReport_Syndicated_V_OnOrder.LeaseReport2.UnitType"
        dbInteger "ColumnWidth" ="2490"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LeaseReport_Syndicated_V_OnOrder.SydCost"
        dbLong "AggregateType" ="0"
    End
    Begin
        dbText "Name" ="OrderedOwnedSyn_LeaseReport.OnOrderCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OwnedCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderedOwnedSyn_LeaseReport.SydCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderedOwnedSyn_LeaseReport_1.OnOrderCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Make"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LeaseReport_Syndicated_V_OnOrder.OnOrderCost"
        dbInteger "ColumnWidth" ="1635"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="0"
    End
    Begin
        dbText "Name" ="SumOfSydCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.ActualDeliveryDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderedOwnedSyn_LeaseReport.*"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderedOwnedSyn_LeaseReport.LeaseReport_Syndicated_V_OnOrder.OrderedYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderedOwnedSyn_LeaseReport.OrderedYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfOnOrderCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OnOrderCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderedOwnedSyn_LeaseReport.OwnedYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfUnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Type"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ModelYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =43
    Top =91
    Right =1268
    Bottom =1021
    Left =-1
    Top =-1
    Right =1193
    Bottom =579
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =118
        Top =57
        Right =464
        Bottom =505
        Top =0
        Name ="vw_Client_FAID_PO_Lease_Sch_Grp_Units"
        Name =""
    End
End
