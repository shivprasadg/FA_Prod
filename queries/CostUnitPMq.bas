dbMemo "SQL" ="SELECT Schedule.SchNo, Locations.LocationName, Units.unitunitnum, Units.unitvin,"
    " IIf(IsNull([unitinserv]),[UnitAcceptDate],[unitinserv]) AS InServDt, Units.unit"
    "yr, Units.unitmake, Units.unitmodel, Units.unittype, SchGrp.BLCD, PMs.PMDate, PM"
    "s.PMno, PMs.MI, PMs.hr, PMs.[Work Type], PMs.[Work Desc]\015\012FROM (((MstrLeas"
    "e INNER JOIN Schedule ON MstrLease.MLID=Schedule.SchMLRef) INNER JOIN SchGrp ON "
    "Schedule.SchID=SchGrp.SGrpSchID) INNER JOIN (Units INNER JOIN Locations ON Units"
    ".UnitGarageLocationRef=Locations.LocationID) ON SchGrp.SGrpID=Units.GroupREF) IN"
    "NER JOIN PMs ON Units.UnitID=PMs.UnitRef\015\012WHERE (((MstrLease.MLOrig)=\"fa\""
    "))\015\012ORDER BY Schedule.SchID, Locations.LocationName, Units.unitunitnum, Sc"
    "hedule.SchID, PMs.PMDate, PMs.PMno, PMs.[Work Desc];\015\012"
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
        dbText "Name" ="Locations.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitunitnum"
        dbInteger "ColumnWidth" ="1590"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InServDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpBlcd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitvin"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unityr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitmake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitmodel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unittype"
        dbInteger "ColumnWidth" ="1800"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PMs.PMDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PMs.PMno"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PMs.[Work Desc]"
        dbInteger "ColumnWidth" ="2790"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PMs.MI"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PMs.hr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PMs.[Work Type]"
        dbLong "AggregateType" ="-1"
    End
End
