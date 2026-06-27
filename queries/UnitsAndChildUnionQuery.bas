dbMemo "SQL" ="SELECT unitID, UnitFinalCost\015\012FROM Units\015\012UNION ALL SELECT unitRef, "
    "ChildFinalCost\015\012FROM ChildParts;\015\012"
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
        dbText "Name" ="unitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitFinalCost"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
