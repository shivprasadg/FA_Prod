dbMemo "SQL" ="SELECT Cost_DecalVendor1.AssetId, Vendors.VendorName AS DecalVendor1, ChildParts"
    ".childFinalCost AS Decal1Cost, ChildParts.ChildInvoiceNum AS Decal1Invoice, Chil"
    "dParts.ChildVendorPytDate AS Decal1Paid\015\012FROM (Cost_DecalVendor1 INNER JOI"
    "N ChildParts ON (Cost_DecalVendor1.AssetId = ChildParts.UnitREF) AND (Cost_Decal"
    "Vendor1.DecalVendor = ChildParts.ChildVendorREF)) INNER JOIN Vendors ON Cost_Dec"
    "alVendor1.DecalVendor = Vendors.VendorID;\015\012"
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
        dbText "Name" ="DecalVendor1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Decal1Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Decal1Invoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Decal1Paid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_DecalVendor1.AssetId"
        dbLong "AggregateType" ="-1"
    End
End
