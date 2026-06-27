Operation =1
Option =0
Where ="(((vw_SixKeys.MLOrig)=\"FA\"))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="Units"
    Name ="Locations"
    Name ="ChildPart_Cost_OEC_summed"
    Name ="SchGrp"
End
Begin OutputColumns
    Expression ="vw_SixKeys.Schedule"
    Expression ="vw_SixKeys.UnitGroup"
    Alias ="UnitNumber"
    Expression ="vw_SixKeys.Unitnum"
    Expression ="vw_SixKeys.Make"
    Expression ="vw_SixKeys.Model"
    Alias ="Type"
    Expression ="[Axle] & \", \" & [UnitSubtype] & \" \" & [UnitType]"
    Alias ="Unit Cost"
    Expression ="CCur([UnitCost])"
    Alias ="Child Cost"
    Expression ="CCur(Nz([ChildPartCost],0))"
    Alias ="Domicile Location"
    Expression ="[locations].[LCity] & \" \" & [locations].[LState] & \" - \" & [locations].[Loca"
        "tionSN]"
    Alias ="Target Delivery"
    Expression ="vw_SixKeys.EsitmatedDeliveryDate"
    Alias ="Spec Letter"
    Expression ="\"\""
    Alias ="Annual Mileage"
    Expression ="IIf([AllowedMilesPeriod] Like \"*Term*\",(([SgprAllow1]/([SchGrp].[LeaseTermPrim"
        "ary]-Nz([SchGrp].[LeaseTermNotice],0)))*12),[SgprAllow1])"
    Alias ="LO Rent"
    Expression ="Units.unitsrent"
    Alias ="Primary Lease Term"
    Expression ="[SchGrp].[LeaseTermPrimary]-Nz([SchGrp].[LeaseTermNotice])"
    Alias ="Synd Rent"
    Expression ="Units.unitrent"
    Alias ="Lease Extend Term"
    Expression ="SchGrp.LeaseExtOption1Term"
    Alias ="Lease Extend Rent"
    Expression ="SchGrp.LeaseExtOption1Rent"
    Expression ="SchGrp.MTMRent"
    Expression ="vw_SixKeys.SchID"
    Expression ="vw_SixKeys.ClientShortName"
    Expression ="vw_SixKeys.clientID"
    Expression ="vw_SixKeys.UnitID"
    Expression ="vw_SixKeys.ClientGroupId"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="Units"
    Expression ="vw_SixKeys.UnitID = Units.UnitID"
    Flag =1
    LeftTable ="Units"
    RightTable ="Locations"
    Expression ="Units.UnitGarageLocationRef = Locations.LocationID"
    Flag =1
    LeftTable ="Units"
    RightTable ="ChildPart_Cost_OEC_summed"
    Expression ="Units.UnitID = ChildPart_Cost_OEC_summed.UnitId"
    Flag =2
    LeftTable ="vw_SixKeys"
    RightTable ="SchGrp"
    Expression ="vw_SixKeys.SGrpID = SchGrp.SGrpID"
    Flag =1
End
Begin OrderBy
    Expression ="vw_SixKeys.Schedule"
    Flag =0
    Expression ="vw_SixKeys.UnitGroup"
    Flag =0
    Expression ="[locations].[LCity] & \" \" & [locations].[LState] & \" - \" & [locations].[Loca"
        "tionSN]"
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
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.SchID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="19"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Make"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.ClientShortName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="20"
    End
    Begin
        dbText "Name" ="UnitNumber"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Model"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Schedule"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.UnitGroup"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="Domicile Location"
        dbInteger "ColumnWidth" ="2970"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="Annual Mileage"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="Target Delivery"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="Synd Rent"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="15"
    End
    Begin
        dbText "Name" ="Unit Cost"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LO Rent"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Spec Letter"
        dbInteger "ColumnOrder" ="13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lease Extend Term"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="16"
    End
    Begin
        dbText "Name" ="Primary Lease Term"
        dbInteger "ColumnOrder" ="14"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Child Cost"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lease Extend Rent"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="17"
    End
    Begin
        dbText "Name" ="SchGrp.MTMRent"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="18"
    End
    Begin
        dbText "Name" ="Type"
        dbInteger "ColumnWidth" ="3330"
        dbInteger "ColumnOrder" ="6"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.clientgroupID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-464
    Top =20
    Right =1019
    Bottom =1001
    Left =-1
    Top =-1
    Right =1465
    Bottom =477
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =134
        Top =13
        Right =397
        Bottom =596
        Top =0
        Name ="vw_Client_FAID_PO_Lease_Sch_Grp_Units"
        Name =""
    End
    Begin
        Left =416
        Top =144
        Right =560
        Bottom =518
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =602
        Top =304
        Right =746
        Bottom =664
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =594
        Top =21
        Right =738
        Bottom =165
        Top =0
        Name ="ChildPart_Cost_OEC_summed"
        Name =""
    End
    Begin
        Left =788
        Top =110
        Right =1045
        Bottom =430
        Top =0
        Name ="SchGrp"
        Name =""
    End
End
