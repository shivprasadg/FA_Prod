Operation =1
Option =0
Where ="(((Units.UnitID)=275946 Or (Units.UnitID)=276013 Or (Units.UnitID)=276014 Or (Un"
    "its.UnitID)=276015 Or (Units.UnitID)=276016 Or (Units.UnitID)=286933 Or (Units.U"
    "nitID)=286932 Or (Units.UnitID)=286931))"
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
    Expression ="vw_SixKeys.VIN"
    Expression ="Units.UnitID"
    Expression ="Units.unitstatus"
    Expression ="UnitVinDecoder.*"
    Alias ="Expr1"
    Expression ="UnitVinDecoder.VEHICLE_TYPE"
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
Begin OrderBy
    Expression ="vw_SixKeys.ClientShortName"
    Flag =0
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
        dbText "Name" ="UnitVinDecoder.TRAILER_TYPE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.BUS_TYPE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.MOTORCYCLE_TYPES"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.GroupREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.MODEL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.FUEL_TYPE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.VIN"
        dbInteger "ColumnWidth" ="2160"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.BODY_TYPE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.CARLINE_CODE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.CHASSIS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.MAKE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.TRIM_LEVEL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.ENGINE_TYPE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.TRANSMISSION"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.DRIVE_LINE_TYPE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.VEHICLE_TYPE"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3255"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.VEHICLE_CLASS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.ASSY_PLANT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.CAB_TYPE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.MODEL_CHANGE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.COUNTRY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.HORSEPOWER"
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
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.MODEL_YEAR"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.UNIT_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.ENGINE_SERIES_CODE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.SERIES"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.STEERING"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.BRAKE_SYSTEM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.FRAME_TYPE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.HYBRID_TYPE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.ENGINE_CODE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.STATUS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.GVWR_CLASS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.GVWR_PROD_RANGE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.WHEEL_BASE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.AXLE_CONFIGURATION"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.WMI_ID_CODE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.PLANT_CODE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.RESTRAINT_SYSTEM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.ENGINE_MODIFIER"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.EMISSIONS_CODE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.DESCRIPTION"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.MANUFACTURER"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.ENGINE_MANUFACTURER"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
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
End
Begin
    State =0
    Left =1
    Top =27
    Right =1869
    Bottom =1033
    Left =-1
    Top =-1
    Right =1836
    Bottom =706
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =458
        Top =360
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
