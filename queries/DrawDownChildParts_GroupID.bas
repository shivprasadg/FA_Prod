dbMemo "SQL" ="SELECT DrawDownAvailChildParts.FAID, DrawDownAvailChildParts.SchID, Vendors.Vend"
    "orID, Vendors.VendorSNm AS Vendor, DrawDownAvailChildParts.ChildMake AS Make, Dr"
    "awDownAvailChildParts.ChildType AS Type, DrawDownAvailChildParts.GroupID\015\012"
    "FROM DrawDownAvailChildParts INNER JOIN Vendors ON DrawDownAvailChildParts.Child"
    "VendorID = Vendors.VendorID\015\012GROUP BY DrawDownAvailChildParts.FAID, DrawDo"
    "wnAvailChildParts.SchID, Vendors.VendorID, Vendors.VendorSNm, DrawDownAvailChild"
    "Parts.ChildMake, DrawDownAvailChildParts.ChildType, DrawDownAvailChildParts.Grou"
    "pID\015\012ORDER BY DrawDownAvailChildParts.ChildMake;\015\012"
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
        dbText "Name" ="DrawDownAvailChildParts.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawDownAvailChildParts.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Make"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawDownAvailChildParts.GroupID"
        dbLong "AggregateType" ="-1"
    End
End
