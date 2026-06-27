Operation =1
Option =0
Begin InputTables
    Name ="vw_Sch_Grp_Units_ChildParts"
    Alias ="DCP"
    Name ="Vendors"
    Name ="DrawnChildInvoices"
End
Begin OutputColumns
    Alias ="Vendor"
    Expression ="Vendors.vendorsnm"
    Alias ="Invoice"
    Expression ="Nz(DCP.ChildInvoiceNum,\"none\")"
    Alias ="InvCost"
    Expression ="Sum(DCP.childCost)"
    Alias ="Parts"
    Expression ="Count(DCP.ChildPartID)"
    Alias ="Drawn"
    Expression ="CDbl(Nz([DrawnChildInvoices].[Parts],0))"
    Expression ="DCP.SchID"
    Expression ="DCP.SGrpID"
End
Begin Joins
    LeftTable ="DCP"
    RightTable ="Vendors"
    Expression ="DCP.ChildVendorID = Vendors.VendorID"
    Flag =2
    LeftTable ="DCP"
    RightTable ="DrawnChildInvoices"
    Expression ="DCP.ChildInvoiceNum = DrawnChildInvoices.Invoice"
    Flag =2
    LeftTable ="DCP"
    RightTable ="DrawnChildInvoices"
    Expression ="DCP.ChildVendorID = DrawnChildInvoices.ChildVendorID"
    Flag =2
End
Begin Groups
    Expression ="Vendors.vendorsnm"
    GroupLevel =0
    Expression ="Nz(DCP.ChildInvoiceNum,\"none\")"
    GroupLevel =0
    Expression ="CDbl(Nz([DrawnChildInvoices].[Parts],0))"
    GroupLevel =0
    Expression ="DCP.SchID"
    GroupLevel =0
    Expression ="DCP.SGrpID"
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
dbMemo "Filter" ="([DrawableChildInvoices0].[Drawn]<>0 Or [DrawableChildInvoices0].[Drawn] IS Null"
    ")"
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
        dbText "Name" ="Parts"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Drawn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DCP.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_vw_Sch_Grp_Units_ChildParts.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DCP.SGrpID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =92
    Top =276
    Right =1207
    Bottom =796
    Left =-1
    Top =-1
    Right =1091
    Bottom =230
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="DCP"
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
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="DrawnChildInvoices"
        Name =""
    End
End
