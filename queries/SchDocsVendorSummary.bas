Operation =1
Option =0
Begin InputTables
    Name ="vw_SixKeys"
    Name ="SchDocsVendorUnion"
End
Begin OutputColumns
    Expression ="vw_SixKeys.Schedule"
    Alias ="Vendor"
    Expression ="[VendorName] & IIf([UnitType] Like \"*tax*\",\" (\" & [UnitType] & \")\",\"\")"
    Alias ="Cost"
    Expression ="Sum(SchDocsVendorUnion.cost)"
    Alias ="VendorCost"
    Expression ="Sum(vw_SixKeys.VendorCost)"
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
    Expression ="Sum(SchDocsVendorUnion.cost)"
    Flag =0
End
Begin Groups
    Expression ="vw_SixKeys.Schedule"
    GroupLevel =0
    Expression ="[VendorName] & IIf([UnitType] Like \"*tax*\",\" (\" & [UnitType] & \")\",\"\")"
    GroupLevel =0
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
        dbText "Name" ="vw_SixKeys.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost"
        dbInteger "ColumnWidth" ="1515"
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
        dbText "Name" ="VendorCost"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =60
    Top =63
    Right =1631
    Bottom =1009
    Left =-1
    Top =-1
    Right =1547
    Bottom =635
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =227
        Bottom =357
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =347
        Top =110
        Right =491
        Bottom =401
        Top =0
        Name ="SchDocsVendorUnion"
        Name =""
    End
End
