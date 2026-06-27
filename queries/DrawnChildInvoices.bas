dbMemo "SQL" ="SELECT Vendors.vendorsnm AS Vendor, nz(DCP.ChildInvoiceNum,\"none\") AS Invoice,"
    " Sum(DCP.childCost) AS TCost, Count(DCP.ChildPartId) AS Parts, DCP.ChildDrawDown"
    "Id, DCP.ChildVendorID\015\012FROM vw_Sch_Grp_Units_ChildParts AS DCP INNER JOIN "
    "Vendors ON DCP.ChildVendorID=Vendors.VendorID\015\012GROUP BY Vendors.vendorsnm,"
    " nz(DCP.ChildInvoiceNum,\"none\"), DCP.ChildDrawDownId, DCP.ChildVendorID\015\012"
    "HAVING (((DCP.ChildDrawDownId) Is Not Null));\015\012"
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
        dbText "Name" ="Parts"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_vw_Sch_Grp_Units_ChildParts.ChildDrawDownId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_vw_Sch_Grp_Units_ChildParts.ChildVendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DCP.ChildVendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DCP.ChildDrawDownId"
        dbLong "AggregateType" ="-1"
    End
End
