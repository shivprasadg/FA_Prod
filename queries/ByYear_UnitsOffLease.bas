Operation =1
Option =0
Where ="(((vw_SixKeys.OffLeaseDate)>='1/1/' & [AskYear] And (vw_SixKeys.OffLeaseDate)<'1"
    "2/31/' & [AskYear]) AND ((vw_SixKeys.MLOrig)=\"FA\"))"
Begin InputTables
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Expression ="vw_SixKeys.ClientShortName"
    Alias ="UnitsOffLease"
    Expression ="Count(vw_SixKeys.AssetID)"
    Expression ="vw_SixKeys.clientID"
    Expression ="vw_SixKeys.ClientGroupId"
    Alias ="FirstOL"
    Expression ="Min(vw_SixKeys.OffLeaseDate)"
    Alias ="LastOL"
    Expression ="Max(vw_SixKeys.OffLeaseDate)"
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
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.ClientShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsOffLease"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastOL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOL"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =165
    Top =27
    Right =1592
    Bottom =1025
    Left =-1
    Top =-1
    Right =1395
    Bottom =680
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =388
        Top =105
        Right =628
        Bottom =495
        Top =0
        Name ="vw_Client_FAID_PO_Lease_Sch_Grp_Units"
        Name =""
    End
End
