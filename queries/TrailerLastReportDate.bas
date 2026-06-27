dbMemo "SQL" ="SELECT TrailerMilesHours.UnitID, Max(TrailerMilesHours.LogActivityDate) AS MaxOf"
    "LogActivityDate\015\012FROM TrailerMilesHours\015\012GROUP BY TrailerMilesHours."
    "UnitID;\015\012"
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
        dbText "Name" ="TrailerMilesHours.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfLogActivityDate"
        dbLong "AggregateType" ="-1"
    End
End
