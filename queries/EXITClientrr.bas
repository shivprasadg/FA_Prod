dbMemo "SQL" ="SELECT vw_SixKeys.clientID, vw_SixKeys.ClientShortName, vw_SixKeys.MLNo, Count(U"
    "nits.UnitID) AS Units, Schedule.SchMLRef AS MLID\015\012FROM ((SchGrp INNER JOIN"
    " Schedule ON SchGrp.SGrpSchID = Schedule.SchID) INNER JOIN Units ON SchGrp.SGrpI"
    "D = Units.GroupREF) INNER JOIN vw_SixKeys ON Schedule.SchMLRef = vw_SixKeys.MLID"
    "\015\012WHERE (((Units.unitstatus)=\"A\" Or (Units.unitstatus)=\"C\") AND ((Date"
    "Diff(\"m\",Date(),CDate(IIf(IsNull([BLCD]),\"1/1/2050\",DateAdd(\"m\",[ExchangeI"
    "TMonth],[BLCD]))))) Between -12 And 25) AND ((vw_SixKeys.MLOrig)=\"fa\"))\015\012"
    "GROUP BY vw_SixKeys.clientID, vw_SixKeys.ClientShortName, vw_SixKeys.MLNo, Sched"
    "ule.SchMLRef\015\012ORDER BY vw_SixKeys.ClientShortName;\015\012"
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
        dbText "Name" ="Units"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MLID"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.clientID"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.ClientShortName"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.MLNo"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
End
