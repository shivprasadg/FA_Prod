Operation =1
Option =0
Begin InputTables
    Name ="Units"
    Name ="Vendors"
    Name ="SchGrp"
End
Begin OutputColumns
    Alias ="SchID"
    Expression ="SchGrp.SGrpSchID"
    Expression ="Units.UnitID"
    Expression ="Vendors.VendorID"
    Alias ="UnitVendor"
    Expression ="Vendors.VendorName"
    Alias ="UnitCost"
    Expression ="Units.Unitfinalcost"
    Expression ="Units.UnitInvoiceNum"
    Expression ="Units.UnitType"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="SchGrp"
    Expression ="Units.GroupREF = SchGrp.SGrpID"
    Flag =2
    LeftTable ="Units"
    RightTable ="Vendors"
    Expression ="Units.VendorREF = Vendors.VendorID"
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
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchID"
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
        dbText "Name" ="Units.UnitType"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =302
    Top =-4
    Right =1873
    Bottom =942
    Left =-1
    Top =-1
    Right =1553
    Bottom =635
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =377
        Top =-5
        Right =521
        Bottom =484
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =180
        Top =343
        Right =324
        Bottom =487
        Top =0
        Name ="Vendors"
        Name =""
    End
    Begin
        Left =158
        Top =96
        Right =302
        Bottom =240
        Top =0
        Name ="SchGrp"
        Name =""
    End
End
