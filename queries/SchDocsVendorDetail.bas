Operation =1
Option =0
Begin InputTables
    Name ="vw_SixKeys"
    Name ="SchDocsVendorUnion"
End
Begin OutputColumns
    Expression ="vw_SixKeys.Schedule"
    Expression ="vw_SixKeys.Unitnum"
    Expression ="vw_SixKeys.VIN"
    Expression ="vw_SixKeys.UnitDescDocVer"
    Alias ="Vendor"
    Expression ="[VendorName] & IIf([UnitType] Like \"*tax*\",\" (\" & [UnitType] & \")\",\"\")"
    Expression ="SchDocsVendorUnion.InvoiceNum"
    Expression ="SchDocsVendorUnion.UnitType"
    Alias ="Cost"
    Expression ="SchDocsVendorUnion.cost"
    Expression ="vw_SixKeys.VendorCost"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="SchDocsVendorUnion"
    Expression ="vw_SixKeys.AssetID = SchDocsVendorUnion.UnitID"
    Flag =1
End
Begin OrderBy
    Expression ="[VendorName] & IIf([UnitType] Like \"*tax*\",\" (\" & [UnitType] & \")\",\"\")"
    Flag =0
    Expression ="SchDocsVendorUnion.cost"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="vw_SixKeys.UnitDescDocVer"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="5400"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost"
        dbInteger "ColumnWidth" ="1350"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor"
        dbInteger "ColumnWidth" ="2805"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchDocsVendorUnion.InvoiceNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.VIN"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Unitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchDocsVendorUnion.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.VendorCost"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-62
    Top =85
    Right =1509
    Bottom =1031
    Left =-1
    Top =-1
    Right =1547
    Bottom =601
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =476
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="SchDocsVendorUnion"
        Name =""
    End
End
