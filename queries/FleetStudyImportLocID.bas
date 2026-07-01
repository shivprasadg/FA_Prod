dbMemo "SQL" ="SELECT FS.VIN, L.LocationID, L.LocationName\015\012FROM FleetStudyImport_Work AS"
    " FS INNER JOIN Locations AS L ON (FS.Location = L.LocationName) AND (FS.ClientGr"
    "oupID = L.ClientGroupID);\015\012"
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
        dbText "Name" ="L.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FS.VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="L.LocationName"
        dbLong "AggregateType" ="-1"
    End
End
