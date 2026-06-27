dbMemo "SQL" ="SELECT ClientGroups.ClientGroupId, 0 AS LocationID, \"Not Required\" AS Location"
    "Name, \"Not Required\" AS LAddrLine1, NULL AS City, NULL AS St, NULL AS Zip\015\012"
    "FROM ClientGroups;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "NoFormat" ="0"
Begin
    Begin
        dbText "Name" ="ClientGroups.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LAddrLine1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Zip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="St"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="City"
        dbLong "AggregateType" ="-1"
    End
End
