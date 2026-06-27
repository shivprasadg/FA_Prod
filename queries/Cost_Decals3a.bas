dbMemo "SQL" ="SELECT DISTINCT First(ChildParts.ChildVendorREF) AS DecalVendor, Max(ChildParts."
    "ChildClass) AS DecalType, ScheduleUnitDocs3.AssetId INTO Cost_DecalVendor3\015\012"
    "FROM (((ScheduleUnitDocs3 INNER JOIN ChildParts ON ScheduleUnitDocs3.AssetId=Chi"
    "ldParts.UnitREF) INNER JOIN Vendors ON ChildParts.ChildVendorREF=Vendors.VendorI"
    "D) LEFT JOIN Cost_DecalVendor1 ON ChildParts.ChildClass=Cost_DecalVendor1.DecalT"
    "ype) LEFT JOIN Cost_DecalVendor2 ON ChildParts.ChildClass=Cost_DecalVendor2.Deca"
    "lType\015\012WHERE (((ChildParts.ChildType) Like \"*Decal*\") AND ((Cost_DecalVe"
    "ndor1.DecalType) Is Null) AND ((Cost_DecalVendor2.DecalType) Is Null))\015\012GR"
    "OUP BY ScheduleUnitDocs3.AssetId\015\012ORDER BY First(ChildParts.ChildVendorREF"
    ");\015\012"
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
        dbText "Name" ="ScheduleUnitDocs3.AssetId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_DecalVendor1.DecalType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_DecalVendor2.DecalType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DecalType"
        dbLong "AggregateType" ="-1"
    End
End
