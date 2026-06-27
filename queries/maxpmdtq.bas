dbMemo "SQL" ="SELECT PMs.UnitRef, Max(PMs.PMDate) AS MaxOfPMDate, Units.GroupREF\015\012FROM P"
    "Ms INNER JOIN Units ON PMs.UnitRef=Units.UnitID\015\012GROUP BY PMs.UnitRef, Uni"
    "ts.GroupREF;\015\012"
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
        dbText "Name" ="MaxOfPMDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PMs.UnitRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.GroupREF"
        dbLong "AggregateType" ="-1"
    End
End
