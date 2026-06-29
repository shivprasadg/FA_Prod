dbMemo "SQL" ="SELECT Vendors.VendorSNm AS Vendor, nz([InvoiceNumber],\"none\") AS Invoice, Sum"
    "(vw_SixKeys.UnitCost) AS InvCost, Count(vw_SixKeys.AssetID) AS Units, CDbl(NZ([D"
    "rawnUnitInvoices].[Units],0)) AS Drawn, vw_SixKeys.SchID\015\012FROM (vw_SixKeys"
    " LEFT JOIN Vendors ON vw_SixKeys.VendorId=Vendors.VendorID) LEFT JOIN DrawnUnitI"
    "nvoices ON (vw_SixKeys.VendorId=DrawnUnitInvoices.VendorID) AND (vw_SixKeys.Invo"
    "iceNumber=DrawnUnitInvoices.Invoice)\015\012GROUP BY Vendors.VendorSNm, nz([Invo"
    "iceNumber],\"none\"), CDbl(NZ([DrawnUnitInvoices].[Units],0)), vw_SixKeys.SchID\015"
    "\012ORDER BY nz([InvoiceNumber],\"none\");\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbMemo "Filter" ="([DrawableUnitInvoices0].[Drawn]<>0 Or [DrawableUnitInvoices0].[Drawn] IS Null)"
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
        dbText "Name" ="Units"
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
End
