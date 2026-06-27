dbMemo "SQL" ="SELECT LookUp_UnitDesc.DescriptionMed AS Description, Avg(Rmkt.SoldPrice) AS Avg"
    "SoldPrice, Count(vw_SixKeys.AssetID) AS Units\015\012FROM (vw_SixKeys INNER JOIN"
    " (Units INNER JOIN Rmkt ON Units.UnitID = Rmkt.UnitRef) ON vw_SixKeys.AssetID = "
    "Units.UnitID) INNER JOIN LookUp_UnitDesc ON vw_SixKeys.AssetID = LookUp_UnitDesc"
    ".AssetID\015\012WHERE (((Rmkt.InvoicePaidDate) Is Null Or (Rmkt.InvoicePaidDate)"
    ">[forms]![acctmenu]![AsOfDate]))\015\012GROUP BY LookUp_UnitDesc.DescriptionMed\015"
    "\012HAVING (((Avg(Rmkt.SoldPrice))>0));\015\012"
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
        dbText "Name" ="Description"
        dbInteger "ColumnWidth" ="4530"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AvgSoldPrice"
        dbInteger "ColumnWidth" ="1860"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
End
