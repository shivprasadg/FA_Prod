dbMemo "SQL" ="SELECT LookUp_UnitDesc.DescriptionMed AS Description, Avg(Rmkt.TargetSale) AS Pr"
    "ojectedSalePrice\015\012FROM Rmkt INNER JOIN LookUp_UnitDesc ON Rmkt.UnitRef = L"
    "ookUp_UnitDesc.AssetID\015\012GROUP BY LookUp_UnitDesc.DescriptionMed\015\012HAV"
    "ING (((Avg(Rmkt.TargetSale)) > 0));\015\012"
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
        dbInteger "ColumnWidth" ="8280"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ProjectedSalePrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AvgOfTargetSale"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.TargetSale"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LookUp_UnitDesc.DescriptionLong"
        dbLong "AggregateType" ="-1"
    End
End
