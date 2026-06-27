Operation =3
Name ="TrackingSheetChangedDatesLog"
Option =0
Where ="((([BuildDate] & [EstimatedDeliveryDate] & [BuildDate]) Is Not Null))"
Begin InputTables
    Name ="OpenVendorDataTemp"
    Name ="Units"
    Name ="vw_SixKeys"
    Name ="OpenVendorTrackingUpdate_BuildDate"
    Name ="OpenVendorTrackingUpdate_EstimatedDeliveryDate"
    Name ="OpenVendorTrackingUpdate_DeliveryDate"
End
Begin OutputColumns
    Name ="UnitID"
    Expression ="Units.UnitID"
    Name ="FAID"
    Expression ="vw_SixKeys.FAID"
    Alias ="SchNo"
    Name ="SchNo"
    Expression ="vw_SixKeys.Schedule"
    Alias ="GrpNo"
    Name ="GrpNo"
    Expression ="vw_SixKeys.UnitGroup"
    Alias ="UnitNum"
    Name ="UnitNum"
    Expression ="Units.UnitUnitNum"
    Alias ="PONum"
    Name ="PONum"
    Expression ="vw_SixKeys.PurchaseOrder"
    Alias ="SubmissionDate"
    Name ="SubmissionDate"
    Expression ="Date()"
    Name ="BuildDate"
    Expression ="OpenVendorTrackingUpdate_BuildDate.BuildDate"
    Name ="EstimatedDeliveryDate"
    Expression ="OpenVendorTrackingUpdate_EstimatedDeliveryDate.EstimatedDeliveryDate"
    Name ="DeliveryDate"
    Expression ="OpenVendorTrackingUpdate_DeliveryDate.DeliveryDate"
    Alias ="PrevBuild"
    Name ="PrevBuildDate"
    Expression ="IIf([PrevBuildDate]=\"1/1/1900\",Null,[PrevBuildDate])"
    Alias ="PrevEstimatedDate"
    Name ="PrevEstimatedDate"
    Expression ="OpenVendorTrackingUpdate_EstimatedDeliveryDate.PrevEstimatedDate"
    Name ="PrevDeliveryDate"
    Expression ="OpenVendorTrackingUpdate_DeliveryDate.PrevDeliveryDate"
End
Begin Joins
    LeftTable ="OpenVendorDataTemp"
    RightTable ="Units"
    Expression ="OpenVendorDataTemp.AssetID = Units.UnitID"
    Flag =1
    LeftTable ="OpenVendorDataTemp"
    RightTable ="vw_SixKeys"
    Expression ="OpenVendorDataTemp.AssetID = vw_SixKeys.AssetID"
    Flag =1
    LeftTable ="Units"
    RightTable ="OpenVendorTrackingUpdate_BuildDate"
    Expression ="Units.UnitID = OpenVendorTrackingUpdate_BuildDate.UnitID"
    Flag =2
    LeftTable ="Units"
    RightTable ="OpenVendorTrackingUpdate_EstimatedDeliveryDate"
    Expression ="Units.UnitID = OpenVendorTrackingUpdate_EstimatedDeliveryDate.UnitID"
    Flag =2
    LeftTable ="Units"
    RightTable ="OpenVendorTrackingUpdate_DeliveryDate"
    Expression ="Units.UnitID = OpenVendorTrackingUpdate_DeliveryDate.UnitID"
    Flag =2
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
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GrpNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PONum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SubmissionDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OpenVendorTrackingUpdate_BuildDate.PrevBuildDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OpenVendorTrackingUpdate_DeliveryDate.PrevDeliveryDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OpenVendorTrackingUpdate_BuildDate.BuildDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OpenVendorTrackingUpdate_DeliveryDate.DeliveryDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OpenVendorTrackingUpdate_EstimatedDeliveryDate.PrevEstimatedDeliveryDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OpenVendorTrackingUpdate_EstimatedDeliveryDate.EstimatedDeliveryDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CDC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EstimatedDeliveryDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PrevBuildDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.BuildDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BuildDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DeliveryDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PrevEstimatedDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PrevBuild"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-21
    Top =4
    Right =1385
    Bottom =1002
    Left =-1
    Top =-1
    Right =1374
    Bottom =527
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =58
        Top =12
        Right =232
        Bottom =349
        Top =0
        Name ="OpenVendorDataTemp"
        Name =""
    End
    Begin
        Left =277
        Top =74
        Right =593
        Bottom =601
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =640
        Top =81
        Right =888
        Bottom =531
        Top =0
        Name ="vw_Client_FAID_PO_Lease_Sch_Grp_Units"
        Name =""
    End
    Begin
        Left =837
        Top =136
        Right =981
        Bottom =280
        Top =0
        Name ="OpenVendorTrackingUpdate_BuildDate"
        Name =""
    End
    Begin
        Left =913
        Top =286
        Right =1057
        Bottom =430
        Top =0
        Name ="OpenVendorTrackingUpdate_EstimatedDeliveryDate"
        Name =""
    End
    Begin
        Left =944
        Top =455
        Right =1088
        Bottom =599
        Top =0
        Name ="OpenVendorTrackingUpdate_DeliveryDate"
        Name =""
    End
End
