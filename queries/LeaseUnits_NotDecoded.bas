Operation =1
Option =0
Where ="(((Units.unitstatus)=\"A\") AND ((UnitVinDecoder.VIN) Is Null))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="Units"
    Name ="UnitVinDecoder"
End
Begin OutputColumns
    Expression ="vw_SixKeys.ClientGroupId"
    Expression ="vw_SixKeys.clientID"
    Expression ="vw_SixKeys.ClientShortName"
    Expression ="vw_SixKeys.MLNo"
    Expression ="vw_SixKeys.Schedule"
    Expression ="vw_SixKeys.UnitGroup"
    Expression ="Units.GroupREF"
    Expression ="Units.UnitID"
    Expression ="Units.UnitVIN"
    Expression ="Units.unitstatus"
    Expression ="vw_SixKeys.UnitGroup"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="UnitVinDecoder"
    Expression ="Units.UnitVIN = UnitVinDecoder.VIN"
    Flag =2
    LeftTable ="vw_SixKeys"
    RightTable ="Units"
    Expression ="vw_SixKeys.AssetID = Units.UnitID"
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
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.GroupREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitstatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.ClientShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1005"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.MODEL_YEAR"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =72
    Top =36
    Right =1595
    Bottom =1042
    Left =-1
    Top =-1
    Right =1491
    Bottom =706
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =458
        Top =439
        Right =602
        Bottom =724
        Top =0
        Name ="UnitVinDecoder"
        Name =""
    End
    Begin
        Left =277
        Top =61
        Right =421
        Bottom =624
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =81
        Top =55
        Right =225
        Bottom =633
        Top =0
        Name ="vw_Client_FAID_PO_Lease_Sch_Grp_Units"
        Name =""
    End
End
