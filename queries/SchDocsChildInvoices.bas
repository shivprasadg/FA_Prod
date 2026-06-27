Operation =6
Option =0
Where ="(((Cost_ChildParts.SchID)=1277))"
Begin InputTables
    Name ="Cost_ChildParts"
    Name ="Vendors"
End
Begin OutputColumns
    Expression ="Cost_ChildParts.AssetID"
    GroupLevel =2
    Alias ="Expr2"
    Expression ="'Vendor :' & [VendorSNm]"
    GroupLevel =1
    Alias ="Expr1"
    Expression ="First(Nz([PartInvoice],\"-\"))"
End
Begin Joins
    LeftTable ="Cost_ChildParts"
    RightTable ="Vendors"
    Expression ="Cost_ChildParts.ChildVendor = Vendors.VendorID"
    Flag =1
End
Begin Groups
    Expression ="Cost_ChildParts.AssetID"
    GroupLevel =2
    Expression ="'Vendor :' & [VendorSNm]"
    GroupLevel =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
Begin
    Begin
        dbText "Name" ="Cost_ChildParts.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unit326280"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Decal Vendor"
        dbInteger "ColumnWidth" ="2085"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PIVOT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Dry Van Vendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Freight Vendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildVendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sales Tax Vendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfAssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unit326277"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unit326278"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AssetID: 326278"
        dbInteger "ColumnWidth" ="1860"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unit326279"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unit326281"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unit326282"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AssetID: 326282"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AssetID: 326277"
        dbInteger "ColumnWidth" ="1860"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AssetID: 326279"
        dbInteger "ColumnWidth" ="1860"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AssetID: 326280"
        dbInteger "ColumnWidth" ="1860"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AssetID: 326281"
        dbInteger "ColumnWidth" ="1860"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor: Advertising Vehicles"
        dbInteger "ColumnWidth" ="3120"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor: Peach State Freightliner"
        dbInteger "ColumnWidth" ="3435"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor: PNC"
        dbInteger "ColumnWidth" ="1605"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfPartCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor :PNC"
        dbInteger "ColumnWidth" ="1635"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfPartInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorSNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor :Advertising Vehicles"
        dbInteger "ColumnWidth" ="3120"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor :Peach State Freightliner"
        dbInteger "ColumnWidth" ="3435"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =88
    Top =146
    Right =1219
    Bottom =1017
    Left =-1
    Top =-1
    Right =1113
    Bottom =543
    Left =0
    Top =0
    ColumnsShown =559
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Cost_ChildParts"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Vendors"
        Name =""
    End
End
