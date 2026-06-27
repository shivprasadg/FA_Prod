Operation =1
Option =0
Where ="(((vw_SixKeys.MLOrig)=\"FA\") AND ((Units.UnitTASent)>='1/1/' & [AskYear] And (U"
    "nits.UnitTASent)<'12/31/' & [AskYear]))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="Units"
End
Begin OutputColumns
    Expression ="vw_SixKeys.ClientShortName"
    Alias ="UnitTitled"
    Expression ="Count(vw_SixKeys.AssetID)"
    Expression ="vw_SixKeys.clientID"
    Expression ="vw_SixKeys.ClientGroupId"
    Alias ="FirstTitled"
    Expression ="Min(Units.UnitTASent)"
    Alias ="LastTitled"
    Expression ="Max(Units.UnitTASent)"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="Units"
    Expression ="vw_SixKeys.UnitID = Units.UnitID"
    Flag =1
End
Begin Groups
    Expression ="vw_SixKeys.ClientShortName"
    GroupLevel =0
    Expression ="vw_SixKeys.clientID"
    GroupLevel =0
    Expression ="vw_SixKeys.ClientGroupId"
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
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.ClientShortName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2115"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitled"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstTitled"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastTitled"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleLog.EntryReasonId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleLog.Reason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTASent"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =164
    Top =65
    Right =1591
    Bottom =1063
    Left =-1
    Top =-1
    Right =1395
    Bottom =612
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =91
        Top =77
        Right =331
        Bottom =520
        Top =0
        Name ="vw_Client_FAID_PO_Lease_Sch_Grp_Units"
        Name =""
    End
    Begin
        Left =441
        Top =419
        Right =585
        Bottom =563
        Top =0
        Name ="Units"
        Name =""
    End
End
