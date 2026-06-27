Operation =1
Option =2
Where ="(((ChildParts.ChildType) Like \"*Decal*\"))"
Having ="(((ChildParts.ChildClass)=\"Misc-2\"))"
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
    Alias ="MiscCost"
    Expression ="ChildParts.childFinalCost"
    Alias ="MiscInvoice"
    Expression ="ChildParts.ChildInvoiceNum"
    Alias ="MiscPaid"
    Expression ="ChildParts.ChildVendorPytDate"
    Alias ="DecalType"
    Expression ="ChildParts.ChildClass"
    Alias ="MiscDraw"
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
        dbText "Name" ="MiscInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MiscCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MiscPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MiscDraw"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AssetId"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =119
    Top =59
    Right =1410
    Bottom =777
    Left =-1
    Top =-1
    Right =1259
    Bottom =92
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =13
        Right =244
        Bottom =359
        Top =0
        Name ="ChildParts"
        Name =""
    End
    Begin
        Left =406
        Top =18
        Right =550
        Bottom =162
        Top =0
        Name ="Vendors"
        Name =""
    End
End
