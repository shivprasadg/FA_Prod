dbMemo "SQL" ="PARAMETERS [enter starting dt] DateTime, [enter ending date] DateTime;\015\012SE"
    "LECT MstrLease.MLOrig, MstrLease.MLNo, Schedule.SchNo, SchGrp.SGrpNO, SchGrp.BLC"
    "D, SchGrp.SGRPNoPD, Schedule.SchFSL, Schedule.SchSLVRec, SchGrp.SGRPINTPD, SchGr"
    "p.SGRPINTRT, SchGrp.SGrpNoSlv, Units.UnitPDStartDate, Units.UnitAcceptDate, Unit"
    "s.UnitInServ, IIf(IsNull([unitpdstartdate]),(IIf(IsNull([unitinserv]),[unitaccep"
    "tdate],[unitinserv])),[unitpdstartdate]) AS PDstart, IIf(((DateDiff(\"d\",[BLCD]"
    ",(IIf(IsNull([unitpdstartdate]),(IIf(IsNull([unitinserv]),[unitacceptdate],[unit"
    "inserv])),[unitpdstartdate]))))*-1)<0 Or [sgrpnopd]=-1,0,(DateDiff(\"d\",[BLCD],"
    "(IIf(IsNull([unitpdstartdate]),(IIf(IsNull([unitinserv]),[unitacceptdate],[uniti"
    "nserv])),[unitpdstartdate]))))*-1) AS PDdays, (DateDiff(\"d\",[BLCD],(IIf(IsNull"
    "([unitpdstartdate]),(IIf(IsNull([unitinserv]),[unitacceptdate],[unitinserv])),[u"
    "nitpdstartdate]))))*-1 AS nock\015\012FROM ((MstrLease INNER JOIN Schedule ON Ms"
    "trLease.MLID = Schedule.SchMLRef) INNER JOIN SchGrp ON Schedule.SchID = SchGrp.S"
    "GrpSchID) INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF\015\012WHERE (((Mst"
    "rLease.MLOrig)=\"fa\") AND ((MstrLease.MLNo) Not Like \"*demo*\") AND ((SchGrp.B"
    "LCD) Is Not Null And (SchGrp.BLCD)>=[enter starting dt] And (SchGrp.BLCD)<=[ente"
    "r ending date]))\015\012ORDER BY MstrLease.MLNo, Schedule.SchID;\015\012"
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
        dbText "Name" ="MstrLease.MLOrig"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.sgrpblcd"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGRPNoPD"
        dbInteger "ColumnOrder" ="17"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchFSL"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchSLVRec"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGRPINTPD"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGRPINTRT"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNoSlv"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPDStartDate"
        dbInteger "ColumnOrder" ="12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitAcceptDate"
        dbInteger "ColumnOrder" ="13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInServ"
        dbInteger "ColumnWidth" ="0"
        dbInteger "ColumnOrder" ="14"
        dbBoolean "ColumnHidden" ="-1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDstart"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDdays"
        dbInteger "ColumnOrder" ="16"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="nock"
        dbInteger "ColumnOrder" ="15"
        dbLong "AggregateType" ="-1"
    End
End
