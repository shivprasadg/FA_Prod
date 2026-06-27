Operation =1
Option =2
Where ="(((ChildParts.ChildType) Like \"*Decal*\"))"
Having ="(((ChildParts.ChildClass) Like \"*Reg*\"))"
Begin InputTables
    Name ="ChildParts"
    Name ="Vendors"
End
Begin OutputColumns
    Alias ="AssetId"
    Expression ="ChildParts.UnitREF"
    Alias ="DecalVendor"
    Expression ="ChildParts.ChildVendorREF"
    Alias ="Name"
    Expression ="Vendors.VendorSNm"
    Alias ="RegDecalCost"
    Expression ="ChildParts.childFinalCost"
    Alias ="RegDecalInvoice"
    Expression ="ChildParts.ChildInvoiceNum"
    Alias ="RegDecalPaid"
    Expression ="ChildParts.ChildVendorPytDate"
    Alias ="DecalType"
    Expression ="ChildParts.ChildClass"
    Alias ="RegDecalDraw"
    Expression ="ChildParts.DrawDownId"
End
Begin Joins
    LeftTable ="ChildParts"
    RightTable ="Vendors"
    Expression ="ChildParts.ChildVendorREF = Vendors.VendorID"
    Flag =1
End
Begin OrderBy
    Expression ="ChildParts.ChildVendorREF"
    Flag =0
End
Begin Groups
    Expression ="ChildParts.UnitREF"
    GroupLevel =0
    Expression ="ChildParts.ChildVendorREF"
    GroupLevel =0
    Expression ="Vendors.VendorSNm"
    GroupLevel =0
    Expression ="ChildParts.childFinalCost"
    GroupLevel =0
    Expression ="ChildParts.ChildInvoiceNum"
    GroupLevel =0
    Expression ="ChildParts.ChildVendorPytDate"
    GroupLevel =0
    Expression ="ChildParts.ChildClass"
    GroupLevel =0
    Expression ="ChildParts.DrawDownId"
    GroupLevel =0
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
        dbText "Name" ="DecalVendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DecalType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RegDecalInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RegDecalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RegDecalPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RegDecalDraw"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AssetId"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =64
    Top =30
    Right =1352
    Bottom =753
    Left =-1
    Top =-1
    Right =1256
    Bottom =245
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =246
        Top =34
        Right =452
        Bottom =380
        Top =0
        Name ="ChildParts"
        Name =""
    End
    Begin
        Left =555
        Top =104
        Right =699
        Bottom =248
        Top =0
        Name ="Vendors"
        Name =""
    End
End
