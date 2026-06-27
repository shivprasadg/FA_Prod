dbMemo "SQL" ="SELECT clientID, ClientShortName, CompanyName, Count(AssetID) AS [Unit Count]\015"
    "\012FROM vw_SixKeys\015\012GROUP BY clientID, ClientShortName, CompanyName\015\012"
    "ORDER BY CompanyName;\015\012"
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
        dbText "Name" ="Unit Count"
        dbLong "AggregateType" ="0"
    End
    Begin
        dbText "Name" ="clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClientShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CompanyName"
        dbLong "AggregateType" ="-1"
    End
End
