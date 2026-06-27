dbMemo "SQL" ="SELECT Units.UnitGarageLocationRef, Min(IIf([UnitStatus]=\"B\",Nz([UnitActualDel"
    "Date],Nz([UnitEstDelDate])),'')) AS FirstDelvery, Max(IIf([UnitStatus]=\"B\",Nz("
    "[UnitActualDelDate],Nz([UnitEstDelDate])),'')) AS LastDelvery\015\012FROM Units\015"
    "\012WHERE (((Units.Unitstatus)=\"B\"))\015\012GROUP BY Units.UnitGarageLocationR"
    "ef;\015\012"
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
        dbText "Name" ="Units.UnitGarageLocationRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstDelvery"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastDelvery"
        dbLong "AggregateType" ="-1"
    End
End
