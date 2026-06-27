Operation =1
Option =2
Where ="(((ChildParts.ChildType) Like \"*Decal*\"))"
Begin InputTables
    Name ="ScheduleUnitDocs3"
    Name ="ChildParts"
End
Begin OutputColumns
    Alias ="DecalVendor"
    Expression ="First(ChildParts.ChildVendorREF)"
    Alias ="AssetId"
    Expression ="ChildParts.UnitREF"
    Alias ="Decal"
    Expression ="ChildParts.childFinalCost"
    Alias ="DecalInvoice"
    Expression ="ChildParts.ChildInvoiceNum"
    Alias ="DecalPaid"
    Expression ="ChildParts.ChildVendorPytDate"
End
Begin Joins
    LeftTable ="ScheduleUnitDocs3"
    RightTable ="ChildParts"
    Expression ="ScheduleUnitDocs3.AssetId = ChildParts.UnitREF"
    Flag =1
End
Begin OrderBy
    Expression ="First(ChildParts.ChildVendorREF)"
    Flag =0
End
Begin Groups
    Expression ="ChildParts.UnitREF"
    GroupLevel =0
    Expression ="ChildParts.childFinalCost"
    GroupLevel =0
    Expression ="ChildParts.ChildInvoiceNum"
    GroupLevel =0
    Expression ="ChildParts.ChildVendorPytDate"
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
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="DecalVendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AssetId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Decal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DecalInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DecalPaid"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =76
    Top =132
    Right =1364
    Bottom =855
    Left =-1
    Top =-1
    Right =1256
    Bottom =415
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =51
        Top =48
        Right =195
        Bottom =192
        Top =0
        Name ="ScheduleUnitDocs3"
        Name =""
    End
    Begin
        Left =246
        Top =34
        Right =452
        Bottom =380
        Top =0
        Name ="ChildParts"
        Name =""
    End
End
