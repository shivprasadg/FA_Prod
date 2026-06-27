Operation =1
Option =0
Where ="(((vw_SixKeys.MLOrig)=\"FA\"))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="Vendors"
End
Begin OutputColumns
    Alias ="Client"
    Expression ="vw_SixKeys.ClientShortName"
    Expression ="Vendors.VendorName"
    Expression ="vw_SixKeys.UnitDescShortVer"
    Alias ="UnitCount"
    Expression ="Count(vw_SixKeys.UnitID)"
    Alias ="Unitcost"
    Expression ="Sum(vw_SixKeys.Unitcost)"
    Alias ="PartsCost"
    Expression ="Sum(vw_SixKeys.PartsCost)"
    Alias ="VendorCost"
    Expression ="Sum(vw_SixKeys.VendorCost)"
    Expression ="Vendors.VendorID"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="Vendors"
    Expression ="vw_SixKeys.VendorId = Vendors.VendorID"
    Flag =1
End
Begin OrderBy
    Expression ="vw_SixKeys.ClientShortName"
    Flag =0
    Expression ="vw_SixKeys.UnitDescShortVer"
    Flag =1
End
Begin Groups
    Expression ="vw_SixKeys.ClientShortName"
    GroupLevel =0
    Expression ="Vendors.VendorName"
    GroupLevel =0
    Expression ="vw_SixKeys.UnitDescShortVer"
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
dbBoolean "TotalsRow" ="-1"
Begin
    Begin
        dbText "Name" ="SumOfChildPartCost"
        dbInteger "ColumnWidth" ="2310"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="0"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Make"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitCount"
        dbLong "AggregateType" ="0"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Model"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfUnitcost"
        dbInteger "ColumnWidth" ="1980"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="0"
    End
    Begin
        dbText "Name" ="vw_SixKeys.PartsCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfPartsCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitDescShortVer"
        dbInteger "ColumnWidth" ="5085"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.ClientShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.VendorCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfVendorCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorCost"
        dbInteger "ColumnWidth" ="1725"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PartsCost"
        dbInteger "ColumnWidth" ="1620"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unitcost"
        dbInteger "ColumnWidth" ="1980"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-1
    Top =9
    Right =1441
    Bottom =1003
    Left =-1
    Top =-1
    Right =1418
    Bottom =666
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =268
        Bottom =619
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =336
        Top =23
        Right =480
        Bottom =167
        Top =0
        Name ="Vendors"
        Name =""
    End
End
