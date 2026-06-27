dbMemo "SQL" ="SELECT PMatchOne.[VIN #], Units.UnitVIN, PMatchOne.[License Plate State]\015\012"
    "FROM Units INNER JOIN PMatchOne ON Units.UnitVIN=PMatchOne.[VIN #];\015\012"
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
        dbText "Name" ="PMatchOne.[VIN #]"
        dbInteger "ColumnWidth" ="2325"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PMatchOne.[License Plate State]"
        dbLong "AggregateType" ="-1"
    End
End
