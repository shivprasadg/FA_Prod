dbMemo "SQL" ="SELECT DISTINCT Max(ChildParts.ChildVendorREF) AS DecalVendor, ScheduleUnitDocs3"
    ".AssetId INTO Cost_DecalVendor4\015\012FROM ((((ScheduleUnitDocs3 INNER JOIN Chi"
    "ldParts ON ScheduleUnitDocs3.AssetId=ChildParts.UnitREF) INNER JOIN Vendors ON C"
    "hildParts.ChildVendorREF=Vendors.VendorID) LEFT JOIN Cost_DecalVendor1 ON ChildP"
    "arts.ChildVendorREF=Cost_DecalVendor1.DecalVendor) LEFT JOIN Cost_DecalVendor2 O"
    "N ChildParts.ChildVendorREF=Cost_DecalVendor2.DecalVendor) LEFT JOIN Cost_DecalV"
    "endor3 ON ChildParts.ChildVendorREF=Cost_DecalVendor3.DecalVendor\015\012WHERE ("
    "((ChildParts.ChildType) Like \"*Decal*\") AND ((Cost_DecalVendor1.DecalVendor) I"
    "s Null) AND ((Cost_DecalVendor2.DecalVendor) Is Null) AND ((Cost_DecalVendor3.De"
    "calVendor) Is Null))\015\012GROUP BY ScheduleUnitDocs3.AssetId\015\012ORDER BY M"
    "ax(ChildParts.ChildVendorREF);\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="DecalVendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildPArts"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.UnitREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AssetId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.childFinalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildInvoiceNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildVendorPytDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Decal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DecalInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DecalPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorSNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfVendorSNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_DecalX.DecalVendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_DecalVendor1.DecalVendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_DecalVendor2.DecalVendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_DecalVendor3.DecalVendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.AssetId"
        dbLong "AggregateType" ="-1"
    End
End
