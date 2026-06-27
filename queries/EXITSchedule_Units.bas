dbMemo "SQL" ="SELECT Units.UnitInServ AS InService, Units.unitstatus AS Status, Units.Axle, Un"
    "its.UNITSRENT AS Rent, SchGrp.BLCD AS BLCD, Round([LeaseTermPrimary]) AS Term, C"
    "Date(IIf(IsNull([BLCD]),\"1/1/2050\",(DateAdd(\"m\",[LeaseTermPrimary],[BLCD])-1"
    "))) AS LeaseEndDate, SchGrp.ExchangeITMonth AS ExchangetITTerm, CDate(IIf(IsNull"
    "([BLCD]),\"1/1/2050\",(DateAdd(\"m\",[ExchangeITMonth],[BLCD])-1))) AS ExchangeI"
    "TDate, CDate(IIf(IsNull([BLCD]),\"1/1/2050\",DateAdd(\"m\",[LeaseTermPrimary],[B"
    "LCD])-90)) AS NoticeDate, IIf(DateDiff(\"m\",Date(),CDate(IIf(IsNull([BLCD]),\"1"
    "/1/2050\",(DateAdd(\"m\",[ExchangeITMonth],[BLCD])-1))))<-3,0,[ExchangeITGroup])"
    " AS ExchangeIT, SchGrp.LeaseExtOption1Term AS ExtOption, Units.UnitSynResid AS E"
    "xpRV, SchGrp.Assignee_Bank AS Assignee, [UnitYr] & \" \" & [UnitMake] & \" \" & "
    "[UnitModel] & \" \" & [UnitSubtype] & \" \" & [UnitType] AS [DESC], DateDiff(\"m"
    "\",Date(),CDate(IIf(IsNull([BLCD]),\"1/1/2050\",DateAdd(\"m\",[ExchangeITMonth],"
    "[BLCD])))) AS MonthsUnitEXIT, DateDiff(\"m\",Date(),CDate(IIf(IsNull([BLCD]),\"1"
    "/1/2050\",DateAdd(\"m\",[LeaseTermPrimary],[BLCD])))) AS MonthsUnitEOL, SchGrp.L"
    "easeExtOption1Term AS ExchangeItMonths\015\012FROM (SchGrp INNER JOIN Schedule O"
    "N SchGrp.SGrpSchID = Schedule.SchID) INNER JOIN Units ON SchGrp.SGrpID = Units.G"
    "roupREF\015\012WHERE (((Units.unitstatus)<>\"S\" And (Units.unitstatus)<>\"KS\")"
    " AND ((DateDiff(\"m\",Date(),CDate(IIf(IsNull([BLCD]),\"1/1/2050\",DateAdd(\"m\""
    ",[ExchangeITMonth],[BLCD]))))) Between -12 And 25));\015\012"
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
        dbText "Name" ="BLCD"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="18"
    End
    Begin
        dbText "Name" ="Term"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="19"
    End
    Begin
        dbText "Name" ="LeaseEndDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="22"
    End
    Begin
        dbText "Name" ="MonthsUnitEOL"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1860"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="31"
    End
    Begin
        dbText "Name" ="MonthsUnitEXIT"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="32"
    End
    Begin
        dbText "Name" ="ExchangetITTerm"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1395"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="33"
    End
    Begin
        dbText "Name" ="DESC"
        dbInteger "ColumnWidth" ="4725"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="34"
    End
    Begin
        dbText "Name" ="NoticeDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="20"
    End
    Begin
        dbText "Name" ="ExchangeIT"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="36"
    End
    Begin
        dbText "Name" ="ExtOption"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="23"
    End
    Begin
        dbText "Name" ="ExchangeItMonths"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="39"
    End
    Begin
        dbText "Name" ="ExchangeITDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1890"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="21"
    End
    Begin
        dbText "Name" ="Assignee"
        dbInteger "ColumnOrder" ="30"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1260"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Status"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="Rent"
        dbInteger "ColumnOrder" ="16"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.Axle"
        dbInteger "ColumnOrder" ="14"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExpRV"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="24"
    End
    Begin
        dbText "Name" ="InService"
        dbLong "AggregateType" ="-1"
    End
End
