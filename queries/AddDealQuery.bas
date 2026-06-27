dbMemo "SQL" ="INSERT INTO Deals ( clientRef, dealShortDesc, dealDate, FAID )\015\012SELECT [Fo"
    "rms]![ClientsView]![clientID] AS clientRef, \"New Transaction Added \" & Now() A"
    "S dealShortDesc, Now() AS dealDate, \"FAID\" AS FAID;\015\012"
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
        dbText "Name" ="Deals.clientRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealShortDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clientRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dealShortDesc"
        dbInteger "ColumnWidth" ="2955"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dealDate"
        dbInteger "ColumnWidth" ="3420"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FAID"
        dbLong "AggregateType" ="-1"
    End
End
