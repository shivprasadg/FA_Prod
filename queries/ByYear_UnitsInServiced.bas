Operation =1
Option =0
Where ="(((vw_SixKeys.InServiceDate)>='1/1/' & [AskYear] And (vw_SixKeys.InServiceDate)<"
    "'12/31/' & [AskYear]) AND ((vw_SixKeys.MLOrig)=\"FA\"))"
Begin InputTables
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Expression ="vw_SixKeys.ClientShortName"
    Alias ="UnitsInServiced"
    Expression ="Count(vw_SixKeys.AssetID)"
    Expression ="vw_SixKeys.clientID"
    Expression ="vw_SixKeys.ClientGroupId"
    Alias ="FirstInServiced"
    Expression ="Min(vw_SixKeys.InServiceDate)"
    Alias ="LastInServiced"
    Expression ="Max(vw_SixKeys.InServiceDate)"
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
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.PurchaseOrderDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitOrdered"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfPurchaseOrderDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.ClientShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.AcceptedDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsOrdered"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.InServiceDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.ActualDeliveryDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FALeased_Unitcount.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.days2Delivery"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FALeased_Unitcount.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsDelivereed"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FALeased_Unitcount.CompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FALeased_Unitcount.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FALeased_Unitcount.UnitCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfAssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsInServiced"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MinOfInServiceDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfInServiceDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstInServiced"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastInServiced"
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
