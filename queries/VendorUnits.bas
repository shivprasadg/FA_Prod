Operation =1
Option =0
Where ="(((Vendors.VendorID)=[Forms]![VendorsForm]![VendorID]))"
Begin InputTables
    Name ="Vendors"
    Name ="Units"
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Alias ="Client"
    Expression ="vw_SixKeys.ClientShortName"
    Expression ="vw_SixKeys.Schedule"
    Alias ="Group"
    Expression ="vw_SixKeys.UnitGroup"
    Expression ="vw_SixKeys.AssetID"
    Expression ="vw_SixKeys.VIN"
    Expression ="vw_SixKeys.MYear"
    Expression ="vw_SixKeys.Make"
    Expression ="vw_SixKeys.Model"
    Expression ="vw_SixKeys.Type"
    Expression ="vw_SixKeys.SubType"
    Expression ="Units.UnitStatus"
    Expression ="Units.UnitFinalCost"
    Alias ="ChildPartCost"
    Expression ="vw_SixKeys.PartsCost"
    Expression ="vw_SixKeys.OECUnit"
    Alias ="InvDate"
    Expression ="vw_SixKeys.UnitInvoiceDate"
    Alias ="Invoice"
    Expression ="vw_SixKeys.UnitInvoiceNum"
    Expression ="vw_SixKeys.VendorCost"
    Alias ="VendorPaid"
    Expression ="Units.UnitVendorPytDate"
    Alias ="VendorTotalCost"
    Expression ="[Units].[VendorCost]+[ChildPartCost]"
    Expression ="Vendors.VendorID"
    Expression ="Vendors.VendorName"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="vw_SixKeys"
    Expression ="Units.UnitID = vw_SixKeys.UnitID"
    Flag =1
    LeftTable ="Vendors"
    RightTable ="vw_SixKeys"
    Expression ="Vendors.VendorID = vw_SixKeys.VendorId"
    Flag =1
End
Begin OrderBy
    Expression ="vw_SixKeys.ClientShortName"
    Flag =0
    Expression ="vw_SixKeys.UnitInvoiceDate"
    Flag =1
    Expression ="Units.UnitVendorPytDate"
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
        dbText "Name" ="Vendors.VendorName"
        dbInteger "ColumnWidth" ="3375"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitFinalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.VIN"
        dbInteger "ColumnWidth" ="2145"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Make"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Model"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.[Type]"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorTotalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.SubType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvDate"
        dbInteger "ColumnWidth" ="1140"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Invoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildPartCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorPaid"
        dbInteger "ColumnWidth" ="2235"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.OECUnit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.VendorCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =49
    Top =15
    Right =1907
    Bottom =724
    Left =-1
    Top =-1
    Right =1834
    Bottom =157
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =691
        Top =45
        Right =835
        Bottom =189
        Top =0
        Name ="Vendors"
        Name =""
    End
    Begin
        Left =422
        Top =4
        Right =667
        Bottom =342
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =39
        Top =21
        Right =215
        Bottom =165
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
End
