Operation =1
Option =0
Where ="(((ChildParts.ChildType) Like \"*Tax*\"))"
Begin InputTables
    Name ="ChildParts"
    Name ="Vendors"
End
Begin OutputColumns
    Alias ="AssetId"
    Expression ="ChildParts.UnitREF"
    Alias ="TaxName"
    Expression ="Vendors.VendorSNm"
    Alias ="FeesTax"
    Expression ="Sum(CCur(Nz([childFinalCost],0)))"
    Alias ="TaxInvoice"
    Expression ="ChildParts.ChildInvoiceNum"
    Alias ="TaxPaid"
    Expression ="ChildParts.ChildVendorPytDate"
    Alias ="TaxDraw"
    Expression ="ChildParts.DrawDownId"
End
Begin Joins
    LeftTable ="ChildParts"
    RightTable ="Vendors"
    Expression ="ChildParts.ChildVendorREF = Vendors.VendorID"
    Flag =1
End
Begin Groups
    Expression ="ChildParts.UnitREF"
    GroupLevel =0
    Expression ="Vendors.VendorSNm"
    GroupLevel =0
    Expression ="ChildParts.ChildInvoiceNum"
    GroupLevel =0
    Expression ="ChildParts.ChildVendorPytDate"
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
        dbText "Name" ="AssetId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FeesTax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TaxInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TaxPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TaxName"
        dbInteger "ColumnWidth" ="1605"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TaxDraw"
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
    Bottom =185
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =179
        Top =-7
        Right =385
        Bottom =243
        Top =0
        Name ="ChildParts"
        Name =""
    End
    Begin
        Left =466
        Top =56
        Right =610
        Bottom =200
        Top =0
        Name ="Vendors"
        Name =""
    End
End
