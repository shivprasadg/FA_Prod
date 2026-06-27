dbMemo "SQL" ="SELECT UnitID, UnitVendor as VendorName, UnitCost as Cost, UnitInvoiceNum as Inv"
    "oiceNum, UnitType, VendorID FROM SchDocsVendorUnit WHERE SchID=3056 UNION ALL Se"
    "lect UnitID, ChildVendor, ChildCost,ChildInvoiceNum, ChildType, ChildVendorID FR"
    "OM SchDocsVendorChild WHERE SchID=3056\015\012ORDER BY UnitId, Cost DESC , UnitT"
    "ype;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
Begin
End
