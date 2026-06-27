Operation =1
Option =0
Begin InputTables
    Name ="Vendors"
    Name ="Units"
    Name ="ChildPartSummaryQuery"
End
Begin OutputColumns
    Expression ="Units.UnitEstDelDate"
    Expression ="Units.UnitUnitNum"
    Expression ="Units.UnitID"
    Expression ="Units.GroupREF"
    Expression ="Units.UnitStatus"
    Expression ="Units.UnitVIN"
    Expression ="Units.UnitInvoiceNum"
    Expression ="Units.UnitFinalCost"
    Alias ="EDIT"
    Expression ="\"Edit\""
    Expression ="Vendors.VendorName"
    Expression ="ChildPartSummaryQuery.SumOfchildFinalCost"
    Alias ="UnitAndChildCost"
    Expression ="IIf([SumOfchildFinalCost] Is Null,[UnitFinalCost],[SumOfchildFinalCost]+[UnitFin"
        "alCost])"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="ChildPartSummaryQuery"
    Expression ="Units.UnitID = ChildPartSummaryQuery.UnitREF"
    Flag =2
    LeftTable ="Vendors"
    RightTable ="Units"
    Expression ="Vendors.VendorID = Units.VendorREF"
    Flag =3
End
Begin OrderBy
    Expression ="Units.UnitUnitNum"
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
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.GroupREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInvoiceNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitFinalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EDIT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildPartSummaryQuery.SumOfchildFinalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitAndChildCost"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2070"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Vendors.VendorName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitEstDelDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1133
    Bottom =686
    Left =-1
    Top =-1
    Right =1101
    Bottom =278
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="Vendors"
        Name =""
    End
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =277
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="ChildPartSummaryQuery"
        Name =""
    End
End
