dbMemo "SQL" ="PARAMETERS [enter start date] DateTime, [enter end date] DateTime;\015\012SELECT"
    " Clients.ClientShNm, MstrLease.MLNo, Schedule.SchNo, SchGrp.Assignee_Bank, SchGr"
    "p.BLCD, SchGrp.ExchangeITGroup, SchGrp.ExchangeITMonth, SchGrp.LeaseTermPrimary,"
    " IIf(IsNull([ExchangeITMonth]) Or [ExchangeITMonth]=0,\"N/A\",DateAdd(\"m\",[Exc"
    "hangeITMonth],[BLCD])) AS Exchangeit, DateAdd(\"m\",[ExchangeITMonth],[BLCD]) AS"
    " EXDT, DateAdd(\"m\",[LeaseTermPrimary],[BLCD])-1 AS Exp, SchGrp.UnitsInGroup, U"
    "nits.UnitUnitNum, Units.UnitVIN, Right([unitvin],6) AS vin4, UCase([unityr] & \""
    " \" & [unitmake] & \" \" & [unitmodel] & \" \" & [unittype]) AS [Desc], Units.un"
    "itstatus, Locations.LocationName, Locations.LCity, Locations.LState\015\012FROM "
    "((((Clients INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client) INNER J"
    "OIN Schedule ON MstrLease.MLID = Schedule.SchMLRef) INNER JOIN SchGrp ON Schedul"
    "e.SchID = SchGrp.SGrpSchID) INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF) "
    "INNER JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID\015\012"
    "WHERE (((SchGrp.BLCD) Is Not Null) AND ((DateAdd(\"m\",[ExchangeITMonth],[BLCD])"
    ")>=[enter start date] And (DateAdd(\"m\",[ExchangeITMonth],[BLCD]))<=[enter end "
    "date]) AND ((Units.unitstatus)=\"a\" Or (Units.unitstatus)=\"c\")) OR (((SchGrp."
    "BLCD) Is Not Null) AND ((DateAdd(\"m\",[ExchangeITMonth],[BLCD]))>=[enter start "
    "date] And (DateAdd(\"m\",[ExchangeITMonth],[BLCD]))<=[enter end date]) AND ((Uni"
    "ts.unitstatus)=\"a\" Or (Units.unitstatus)=\"c\")) OR (((SchGrp.BLCD) Is Not Nul"
    "l) AND ((DateAdd(\"m\",[LeaseTermPrimary],[BLCD])-1)>=[enter start date] And (Da"
    "teAdd(\"m\",[LeaseTermPrimary],[BLCD])-1)<=[enter end date]) AND ((Units.unitsta"
    "tus)=\"a\" Or (Units.unitstatus)=\"c\")) OR (((SchGrp.BLCD) Is Not Null) AND ((D"
    "ateAdd(\"m\",[LeaseTermPrimary],[BLCD])-1)>=[enter start date] And (DateAdd(\"m\""
    ",[LeaseTermPrimary],[BLCD])-1)<=[enter end date]))\015\012ORDER BY Units.UnitUni"
    "tNum, Right([unitvin],6);\015\012"
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
        dbText "Name" ="Clients.ClientShNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Exchangeit"
        dbLong "AggregateType" ="-1"
        dbBoolean "ColumnHidden" ="-1"
    End
    Begin
        dbText "Name" ="Exp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Desc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vin4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitstatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EXDT"
        dbLong "AggregateType" ="-1"
    End
End
