Operation =1
Option =0
Begin InputTables
    Name ="Units"
    Name ="ChildParts"
    Name ="SchGrp"
    Name ="Vendors"
End
Begin OutputColumns
    Alias ="SchID"
    Expression ="SchGrp.SGrpSchID"
    Expression ="Units.UnitID"
    Alias ="ChildVendorID"
    Expression ="Vendors.VendorID"
    Alias ="ChildVendor"
    Expression ="Vendors.VendorName"
    Alias ="ChildCost"
    Expression ="ChildParts.childFinalCost"
    Expression ="ChildParts.ChildInvoiceNum"
    Expression ="ChildParts.ChildType"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="ChildParts"
    Expression ="Units.UnitID = ChildParts.UnitREF"
    Flag =2
    LeftTable ="Units"
    RightTable ="SchGrp"
    Expression ="Units.GroupREF = SchGrp.SGrpID"
    Flag =2
    LeftTable ="ChildParts"
    RightTable ="Vendors"
    Expression ="ChildParts.ChildVendorREF = Vendors.VendorID"
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
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildVendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildVendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildInvoiceNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInvoiceNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorName"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =90
    Top =34
    Right =1431
    Bottom =963
    Left =-1
    Top =-1
    Right =1323
    Bottom =584
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =496
        Top =0
        Name ="ChildParts"
        Name =""
    End
    Begin
        Left =50
        Top =188
        Right =194
        Bottom =332
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =434
        Top =199
        Right =578
        Bottom =343
        Top =0
        Name ="Vendors"
        Name =""
    End
End
