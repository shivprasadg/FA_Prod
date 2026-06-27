Operation =1
Option =0
Where ="(((vw_SixKeys.PurchaseOrderDate)>='1/1/' & [AskYear] And (vw_SixKeys.PurchaseOrd"
    "erDate)<'12/31/' & [AskYear]) AND ((vw_SixKeys.MLOrig)=\"FA\"))"
Begin InputTables
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Expression ="vw_SixKeys.ClientShortName"
    Alias ="UnitOrdered"
    Expression ="Count(vw_SixKeys.AssetID)"
    Expression ="vw_SixKeys.clientID"
    Expression ="vw_SixKeys.ClientGroupId"
    Alias ="FirstPO"
    Expression ="Min(vw_SixKeys.PurchaseOrderDate)"
    Alias ="LastPO"
    Expression ="Max(vw_SixKeys.PurchaseOrderDate)"
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
        dbText "Name" ="UnitOrdered"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.ClientShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.PurchaseOrderDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfPurchaseOrderDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MinOfPurchaseOrderDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstPO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastPO"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =88
    Top =24
    Right =1515
    Bottom =1022
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
