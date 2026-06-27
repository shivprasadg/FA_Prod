Operation =1
Option =16
RowCount ="10"
Where ="(((SchGrp.FundDate) Between #3/1/2016# And #2/28/2018#))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="Vendors"
    Name ="UnitTotalOEC"
    Name ="SchGrp"
End
Begin OutputColumns
    Alias ="Vendor"
    Expression ="Vendors.VendorSNm"
    Alias ="SumOfUnitTotalCost"
    Expression ="Sum(UnitTotalOEC.UnitTotalCost)"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="Vendors"
    Expression ="vw_SixKeys.VendorId = Vendors.VendorID"
    Flag =2
    LeftTable ="vw_SixKeys"
    RightTable ="UnitTotalOEC"
    Expression ="vw_SixKeys.AssetID = UnitTotalOEC.AssetID"
    Flag =2
    LeftTable ="vw_SixKeys"
    RightTable ="SchGrp"
    Expression ="vw_SixKeys.SGrpID = SchGrp.SGrpID"
    Flag =1
End
Begin OrderBy
    Expression ="Sum(UnitTotalOEC.UnitTotalCost)"
    Flag =1
End
Begin Groups
    Expression ="Vendors.VendorSNm"
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
        dbText "Name" ="vw_SixKeys.UnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.FundedDate"
        dbInteger "ColumnWidth" ="1560"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTotalOEC.UnitTotalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfUnitTotalCost"
        dbInteger "ColumnWidth" ="1935"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost"
        dbInteger "ColumnWidth" ="1620"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfUnitCost"
        dbInteger "ColumnWidth" ="2520"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorSNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor"
        dbInteger "ColumnWidth" ="2880"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.FundDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MinOfFundDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1550
    Bottom =986
    Left =-1
    Top =-1
    Right =1532
    Bottom =669
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =479
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =256
        Top =260
        Right =400
        Bottom =404
        Top =0
        Name ="Vendors"
        Name =""
    End
    Begin
        Left =365
        Top =84
        Right =509
        Bottom =228
        Top =0
        Name ="UnitTotalOEC"
        Name =""
    End
    Begin
        Left =553
        Top =213
        Right =697
        Bottom =604
        Top =0
        Name ="SchGrp"
        Name =""
    End
End
