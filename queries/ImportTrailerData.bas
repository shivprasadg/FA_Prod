dbMemo "SQL" ="SELECT Units.UnitID, Units.UnitUnitNum, TrailerDataImport.MILRDDT, TrailerLastMa"
    "xMilesHours.MilesLTD AS LastMilesLTD, TrailerDataImport.MILES, [MILES]-[LastMile"
    "sLTD] AS NewMiles, TrailerDataImport.RFRHR, TrailerLastMaxMilesHours.HoursLTD AS"
    " LastHoursLTD, [RFRHR]-[LastHoursLTD] AS NewHours, 1 AS OBCDataStaus, TrailerLas"
    "tMaxMilesHours.LogActivityDate AS LastLogActivityDate, CDate(DatePart('m',[MILRD"
    "DT]) & \"/1/\" & DatePart('yyyy',[MILRDDT])) AS NewLogActivityDate, DateDiff('m'"
    ",[LastLogActivityDate],[NewLogActivityDate]) AS MonthsBetween, CInt([NewMiles]/["
    "MonthsBetween]) AS AvgMiles\015\012FROM (Units INNER JOIN TrailerDataImport ON U"
    "nits.UnitVIN = TrailerDataImport.VIN) LEFT JOIN TrailerLastMaxMilesHours ON Unit"
    "s.UnitID = TrailerLastMaxMilesHours.UnitID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="TrailerDataImport.MILES"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TrailerDataImport.RFRHR"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TrailerDataImport.MILRDDT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastMilesLTD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastHoursLTD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastLogActivityDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCDataStaus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewMiles"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewLogActivityDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MonthsBetween"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewHours"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AvgMiles"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
