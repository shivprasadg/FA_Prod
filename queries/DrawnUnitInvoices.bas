Operation =1
Option =0
Having ="(((Units.DrawDownId) Is Not Null))"
Begin InputTables
    Name ="Units"
    Name ="Vendors"
End
Begin OutputColumns
    Expression ="Units.DrawDownId"
    Alias ="Vendor"
    Expression ="Vendors.vendorsnm"
    Alias ="Invoice"
    Expression ="Nz([UnitInvoiceNum],\"none\")"
    Alias ="Units"
    Expression ="Count(Units.UnitId)"
    Alias ="TCost"
    Expression ="Sum(Units.UnitFinalCost)"
    Expression ="Vendors.VendorID"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="Vendors"
    Expression ="Units.VendorREF = Vendors.VendorID"
    Flag =1
End
Begin Groups
    Expression ="Units.DrawDownId"
    GroupLevel =0
    Expression ="Vendors.vendorsnm"
    GroupLevel =0
    Expression ="Nz([UnitInvoiceNum],\"none\")"
    GroupLevel =0
    Expression ="Vendors.VendorID"
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
        dbText "Name" ="Vendor"
        dbInteger "ColumnWidth" ="1800"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Invoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.DrawDownId"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =947
    Bottom =560
    Left =-1
    Top =-1
    Right =915
    Bottom =169
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =250
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="Vendors"
        Name =""
    End
End
