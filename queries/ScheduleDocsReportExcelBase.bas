Operation =1
Option =0
Where ="(((vw_SixKeys.MLOrig)=\"FA\"))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="SchGrp"
    Name ="Units"
    Name ="Locations"
    Name ="Locations"
    Alias ="Locations_1"
    Name ="vw_LocationTaxes"
    Name ="Locations"
    Alias ="Locations_2"
    Name ="Locations"
    Alias ="Locations_3"
End
Begin OutputColumns
    Expression ="vw_SixKeys.Schedule"
    Expression ="vw_SixKeys.UnitGroup"
    Alias ="Lessee"
    Expression ="vw_SixKeys.CompanyName"
    Expression ="vw_SixKeys.FAID"
    Alias ="Units"
    Expression ="vw_SixKeys.UnitID"
    Alias ="UnitDesc"
    Expression ="Nz([LegalDescription],[UnitDescDocVer])"
    Alias ="UnitCost"
    Expression ="vw_SixKeys.UnitCost"
    Alias ="Child Cost"
    Expression ="Nz(DSum(\"ChildPartCost\",\"ChildPart_Cost_OEC_Unit\",\"UnitId=\" & [Units].[Uni"
        "tID]),0)"
    Alias ="OEC"
    Expression ="[UnitCost]+Nz(DSum(\"ChildPartCost\",\"ChildPart_Cost_OEC_Unit\",\"UnitId=\" & ["
        "Units].[UnitID]),0)"
    Alias ="InvoicePaid"
    Expression ="Units.UnitVendorPytDate"
    Expression ="SchGrp.LeaseTermPrimary"
    Expression ="SchGrp.LeasePaidAdvArr"
    Alias ="Allowed Miles"
    Expression ="SchGrp.SGprAllow1"
    Alias ="MileChrg"
    Expression ="SchGrp.SGrpov1"
    Alias ="MilesThreshold"
    Expression ="SchGrp.SGprAllow2"
    Alias ="MileChrg After"
    Expression ="SchGrp.SGrpov2"
    Alias ="Term Year"
    Expression ="SchGrp.SGprSalesIndYr"
    Alias ="Rate"
    Expression ="SchGrp.SgrpSalesIndRt"
    Alias ="Date"
    Expression ="SchGrp.SGrpSalesIndDt"
    Alias ="Neg BLCD"
    Expression ="DLookUp(\"StartsOn\",\"Lookup_LeaseStartsOn\",\"ID=\" & Nz([SgrpCommInd],10))"
    Alias ="PerDiem Start"
    Expression ="DLookUp(\"StartsOn\",\"Lookup_LeaseStartsOn\",\"ID=\" & Nz([PerDiemStartsOn],10)"
        ")"
    Expression ="SchGrp.BLCD_Projected"
    Alias ="Assignee"
    Expression ="vw_SixKeys.BankShortName"
    Alias ="Spread"
    Expression ="[SgrpSalesYield]-[SgrpSalesIndRt]"
    Alias ="Yield"
    Expression ="SchGrp.SgrpSalesYield"
    Alias ="Adjustment"
    Expression ="SchGrp.BPChgLO_Up"
    Expression ="vw_SixKeys.SchID"
    Alias ="Domicle"
    Expression ="IIf(IsNull([Locations_1].[LAddrLine1]),Null,[Locations_1].[LAddrLine1] & \" \") "
        "& IIf(IsNull([Locations_1].[LAddrLine2]),Null,[Locations_1].[LAddrLine2] & \" \""
        ") & [Locations_1].[LCity] & \" \" & [Locations_1].[LState] & \" \" & [Locations_"
        "1].[LZip] & \" (\" & [Locations_1].[LCounty] & \" County)\""
    Alias ="DeliveryLocation"
    Expression ="IIf(IsNull([Locations].[LAddrLine1]),Null,[Locations].[LAddrLine1] & \" \") & II"
        "f(IsNull([Locations].[LAddrLine2]),Null,[Locations].[LAddrLine2] & \" \") & [Loc"
        "ations].[LCity] & \" \" & [Locations].[LState] & \" \" & [Locations].[LZip] & \""
        " (\" & [Locations].[LCounty] & \" County)\""
    Alias ="SalesTax"
    Expression ="[StateSales]+[CountySales]+[CitySales]"
    Alias ="IRP"
    Expression ="Units.UnitIRP"
    Alias ="RentReceipt"
    Expression ="vw_LocationTaxes.StateRReceipt"
    Alias ="AvalException"
    Expression ="Nz([StateException],Nz([CountyException],[CityException]))"
    Alias ="Billing"
    Expression ="IIf(IsNull([Locations_1].[LAddrLine1]),Null,[Locations_2].[LAddrLine1] & \" \") "
        "& IIf(IsNull([Locations_2].[LAddrLine2]),Null,[Locations_2].[LAddrLine2] & \" \""
        ") & [Locations_2].[LCity] & \" \" & [Locations_2].[LState] & \" \" & [Locations_"
        "2].[LZip]"
    Expression ="SchGrp.BillingInstructions"
    Expression ="SchGrp.TaxInstructions"
    Alias ="Title"
    Expression ="IIf(IsNull([Locations_1].[LAddrLine1]),Null,[Locations_3].[LAddrLine1] & \" \") "
        "& IIf(IsNull([Locations_3].[LAddrLine2]),Null,[Locations_3].[LAddrLine2] & \" \""
        ") & [Locations_3].[LCity] & \" \" & [Locations_3].[LState] & \" \" & [Locations_"
        "3].[LZip]"
    Alias ="SchRent"
    Expression ="Units.UnitSRent"
    Alias ="ExtdRent"
    Expression ="Units.FirstExtRent"
    Alias ="FATitleLogNum"
    Expression ="Units.AssetID_FATitleLog"
    Expression ="vw_SixKeys.VendorCost"
    Expression ="vw_SixKeys.VIN"
    Expression ="vw_SixKeys.UnitNum"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="SchGrp"
    Expression ="vw_SixKeys.SGrpID = SchGrp.SGrpID"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="Units"
    Expression ="vw_SixKeys.UnitID = Units.UnitID"
    Flag =1
    LeftTable ="Units"
    RightTable ="Locations"
    Expression ="Units.UnitDeliveryLocationRef = Locations.LocationID"
    Flag =2
    LeftTable ="Units"
    RightTable ="Locations_1"
    Expression ="Units.UnitGarageLocationRef = Locations_1.LocationID"
    Flag =2
    LeftTable ="Locations_1"
    RightTable ="vw_LocationTaxes"
    Expression ="Locations_1.LocationID = vw_LocationTaxes.LocationID"
    Flag =2
    LeftTable ="SchGrp"
    RightTable ="Locations_2"
    Expression ="SchGrp.BillToAddressID = Locations_2.LocationID"
    Flag =2
    LeftTable ="Units"
    RightTable ="Locations_3"
    Expression ="Units.UnitTitleLocationRef = Locations_3.LocationID"
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
Begin
    Begin
        dbText "Name" ="vw_SixKeys.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitDesc"
        dbInteger "ColumnWidth" ="4500"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OEC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LeaseTermPrimary"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Assignee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lessee"
        dbInteger "ColumnWidth" ="3120"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MileChrg After"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="0"
    End
    Begin
        dbText "Name" ="DeliveryLocation"
        dbInteger "ColumnWidth" ="1455"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LeasePaidAdvArr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Allowed Miles"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MilesThreshold"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MileChrg"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Term year"
        dbInteger "ColumnWidth" ="1350"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Neg Blcd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PerDiem Start"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2055"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Spread"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BLCD_Projected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Yield"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Adjustment"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesTax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AvalException"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoicePaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Child Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RentReceipt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Domicle"
        dbInteger "ColumnWidth" ="1455"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Billing"
        dbInteger "ColumnWidth" ="1455"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.TaxInstructions"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BillingInstructions"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="IRP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Title"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.FirstExtRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtdRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitSRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1026"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.AssetID_FATitleLog"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FATitleLogNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.vendorCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitNum"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =97
    Top =155
    Right =1584
    Bottom =1035
    Left =-1
    Top =-1
    Right =1463
    Bottom =321
    Left =120
    Top =0
    ColumnsShown =539
    Begin
        Left =12
        Top =12
        Right =129
        Bottom =566
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =457
        Top =11
        Right =686
        Bottom =252
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =247
        Top =107
        Right =391
        Bottom =494
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =447
        Top =193
        Right =591
        Bottom =337
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =624
        Top =315
        Right =768
        Bottom =459
        Top =0
        Name ="Locations_1"
        Name =""
    End
    Begin
        Left =829
        Top =322
        Right =973
        Bottom =466
        Top =0
        Name ="vw_LocationTaxes"
        Name =""
    End
    Begin
        Left =807
        Top =109
        Right =951
        Bottom =253
        Top =0
        Name ="Locations_2"
        Name =""
    End
    Begin
        Left =1020
        Top =323
        Right =1164
        Bottom =467
        Top =0
        Name ="Locations_3"
        Name =""
    End
End
