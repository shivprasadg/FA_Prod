Operation =1
Option =0
Where ="(((ChildParts.DrawDownId) Is Null) AND ((vw_SixKeys.MLOrig)=\"FA\"))"
Begin InputTables
    Name ="Vendors"
    Name ="vw_SixKeys"
    Name ="ChildParts"
    Name ="DrawDowns"
End
Begin OutputColumns
    Expression ="ChildParts.ChildPartID"
    Expression ="vw_SixKeys.AssetID"
    Expression ="vw_SixKeys.VIN"
    Expression ="vw_SixKeys.Unitnum"
    Alias ="CInvoice"
    Expression ="ChildParts.ChildInvoiceNum"
    Expression ="ChildParts.ChildType"
    Alias ="PCost"
    Expression ="ChildParts.childFinalCost"
    Expression ="ChildParts.DrawDownId"
    Alias ="CVendor"
    Expression ="Vendors.vendorsnm"
    Alias ="Client"
    Expression ="vw_SixKeys.ClientShortName"
    Expression ="vw_SixKeys.MLNo"
    Expression ="vw_SixKeys.Schedule"
    Expression ="vw_SixKeys.UnitGroup"
    Alias ="RequestNumber"
    Expression ="DrawDowns.DrawRequestNumber"
    Expression ="DrawDowns.DrawName"
    Alias ="ChildDesc"
    Expression ="[ChildMake] & \" \" & [ChildClass] & \" \" & [ChildType]"
    Expression ="Vendors.VendorID"
    Expression ="vw_SixKeys.clientID"
    Expression ="vw_SixKeys.MLID"
    Expression ="vw_SixKeys.SchID"
    Alias ="GroupID"
    Expression ="vw_SixKeys.SGrpID"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="ChildParts"
    Expression ="vw_SixKeys.UnitID = ChildParts.UnitREF"
    Flag =1
    LeftTable ="Vendors"
    RightTable ="ChildParts"
    Expression ="Vendors.VendorID = ChildParts.ChildVendorREF"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="DrawDowns"
    Expression ="vw_SixKeys.DrawDownId = DrawDowns.DrawDownId"
    Flag =2
End
Begin OrderBy
    Expression ="Vendors.vendorsnm"
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
        dbText "Name" ="vw_SixKeys.clientID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="18"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1785"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="19"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Schedule"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="vw_SixKeys.SchID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="20"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitGroup"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="Client"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="CVendor"
        dbInteger "ColumnWidth" ="2085"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="ChildParts.DrawDownId"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1620"
        dbInteger "ColumnOrder" ="8"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="ChildParts.ChildPartID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="CInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="Vendors.VendorID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="17"
    End
    Begin
        dbText "Name" ="RequestNumber"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="ChildDesc"
        dbInteger "ColumnWidth" ="3870"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="16"
    End
    Begin
        dbText "Name" ="DrawDowns.DrawName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2070"
        dbInteger "ColumnOrder" ="15"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="vw_SixKeys.assetId"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="vw_SixKeys.VIN"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.unitnum"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildType"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PCost"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupID"
        dbInteger "ColumnOrder" ="21"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =65
    Top =183
    Right =1434
    Bottom =1022
    Left =-1
    Top =-1
    Right =1345
    Bottom =474
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =572
        Top =-4
        Right =716
        Bottom =140
        Top =0
        Name ="Vendors"
        Name =""
    End
    Begin
        Left =592
        Top =154
        Right =782
        Bottom =431
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =227
        Top =104
        Right =465
        Bottom =406
        Top =0
        Name ="ChildParts"
        Name =""
    End
    Begin
        Left =821
        Top =204
        Right =965
        Bottom =424
        Top =0
        Name ="DrawDowns"
        Name =""
    End
End
