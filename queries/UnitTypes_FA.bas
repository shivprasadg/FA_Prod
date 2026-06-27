Operation =1
Option =0
Having ="(((vw_SixKeys.MLOrig)=\"FA\"))"
Begin InputTables
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Expression ="vw_SixKeys.MLOrig"
    Expression ="vw_SixKeys.Type"
    Alias ="CountOfAssetID"
    Expression ="Count(vw_SixKeys.AssetID)"
End
Begin Groups
    Expression ="vw_SixKeys.MLOrig"
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
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfAssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1474
    Bottom =1046
    Left =-1
    Top =-1
    Right =1442
    Bottom =706
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =70
        Top =74
        Right =463
        Bottom =345
        Top =0
        Name ="vw_Client_FAID_PO_Lease_Sch_Grp_Units"
        Name =""
    End
End
