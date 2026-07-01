dbMemo "SQL" ="SELECT LeaseOptionID, UO.UnitID, StartOrPurposedDate AS XStart, EndOrActualDate "
    "AS XEnd, DatePosted, OptionTypeID\015\012FROM UnitExitOption AS UO INNER JOIN (S"
    "ELECT Ex.UnitID, Max(Ex.EndOrActualDate) AS MaxEndDate FROM UnitExitOption AS Ex"
    " WHERE OptionTypeID > 2 GROUP BY Ex.UnitID)  AS X ON (X.MaxEndDate = UO.EndOrAct"
    "ualDate) AND (UO.UnitID = X.UnitID);\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="OptionTypeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LeaseOptionID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UO.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="XStart"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="XEnd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DatePosted"
        dbLong "AggregateType" ="-1"
    End
End
