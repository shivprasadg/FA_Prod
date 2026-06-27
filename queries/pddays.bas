dbMemo "SQL" ="SELECT Clients.ClientShNm, Schedule.SchNo, SchGrp.SGrpNO, SchGrp.BLCD, Units.uni"
    "tinserv, Units.unitacceptdate, Units.UnitPDStartDate, SchGrp.sgrpnopd, IIf(IsNul"
    "l([unitpdstartdate]),IIf(IsNull([unitinserv]),[unitacceptdate],[unitinserv]),[un"
    "itpdstartdate]) AS PDStart, DateDiff(\"d\",IIf(IsNull([unitpdstartdate]),IIf(IsN"
    "ull([unitinserv]),[unitacceptdate],[unitinserv]),[unitpdstartdate]),[unitacceptd"
    "ate]) AS DaysafteAccpt, IIf(DateDiff(\"d\",IIf(IsNull([unitpdstartdate]),IIf(IsN"
    "ull([unitinserv]),[unitacceptdate],[unitinserv]),[unitpdstartdate]),[BLCD])-1<0,"
    "0,DateDiff(\"d\",IIf(IsNull([unitpdstartdate]),IIf(IsNull([unitinserv]),[unitacc"
    "eptdate],[unitinserv]),[unitpdstartdate]),[BLCD])-1) AS PDdays\015\012FROM Clien"
    "ts INNER JOIN (((MstrLease INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchM"
    "LRef) INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID) INNER JOIN Units O"
    "N SchGrp.SGrpID = Units.GroupREF) ON Clients.clientID = MstrLease.Client\015\012"
    "WHERE (((Clients.ClientShNm) Not Like \"Demo*\") AND ((SchGrp.BLCD) Is Not Null "
    "And (SchGrp.BLCD)>#12/31/2010# And (SchGrp.BLCD)<#1/1/2012#) AND ((Schedule.SchN"
    "o)<>\"1-ext\") AND ((MstrLease.MLOrig)=\"fa\"))\015\012ORDER BY Clients.ClientSh"
    "Nm, Schedule.SchID, Schedule.SchNo, SchGrp.SGrpNO;\015\012"
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
        dbInteger "ColumnWidth" ="1530"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbInteger "ColumnWidth" ="1830"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.sgrpBlcd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitinserv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitacceptdate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPDStartDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.sgrpnopd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDStart"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DaysafteAccpt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDdays"
        dbLong "AggregateType" ="-1"
    End
End
