dbMemo "SQL" ="SELECT Clients.clientID, MstrLease.MLOrig, MstrLease.MLID, MstrLease.MLNo, Sched"
    "ule.SchFAID, Schedule.SCHNo1, Schedule.SchID, Schedule.SchFSL, Schedule.SchMLRef"
    ", Schedule.SchNo, Schedule.SchAlias, Schedule.SchOut, Schedule.SchIn, Schedule.S"
    "chsigned, Schedule.SchSigner, Schedule.SchSigner2, Schedule.SchSigner3, Schedule"
    ".SchAssDt, Schedule.SchTitleAssignee, Schedule.SchTitleToTrust, Schedule.SchShor"
    "tDesc, Schedule.SchFAID, Schedule.SCHINSAMT, Schedule.SCHINREC, Schedule.SCHINSE"
    "XP, Schedule.SchTitltoFA, Schedule.SchTitletoClient, Schedule.LienHolder, Schedu"
    "le.LienHolderNm, MstrLease.MLLessor, MstrLease.Client, Clients.ClientShNm, \"Edi"
    "t\" AS Edit, Schedule.SchAH, IIf(IsNull([Assignee_Bank]),Null,[Assignee_Bank]) A"
    "S Bassign, IIf(IsNull([SumOfOrigEquipCost]),Null,[SumOfOrigEquipCost]) AS ECost,"
    " Max(Units.UnitEstDelDate) AS MaxOfUnitEstDelDate, IIf(IsNull([SgrpPlb]),No,[Sgr"
    "pPlb]) AS PLB, SchGrp.Assigned_Indicator\015\012FROM (((Clients INNER JOIN MstrL"
    "ease ON Clients.clientID = MstrLease.Client) INNER JOIN (Schedule LEFT JOIN SchE"
    "quipCostq ON Schedule.SchID = SchEquipCostq.SchID) ON MstrLease.MLID = Schedule."
    "SchMLRef) LEFT JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID) LEFT JOIN Units"
    " ON SchGrp.SGrpID = Units.GroupREF\015\012GROUP BY Clients.clientID, MstrLease.M"
    "LOrig, MstrLease.MLID, MstrLease.MLNo, Schedule.SCHNo1, Schedule.SchID, Schedule"
    ".SchFSL, Schedule.SchMLRef, Schedule.SchNo, Schedule.SchAlias, Schedule.SchOut, "
    "Schedule.SchIn, Schedule.Schsigned, Schedule.SchSigner, Schedule.SchSigner2, Sch"
    "edule.SchSigner3, Schedule.SchAssDt, Schedule.SchTitleAssignee, Schedule.SchTitl"
    "eToTrust, Schedule.SchShortDesc, Schedule.SchFAID, Schedule.SCHINSAMT, Schedule."
    "SCHINREC, Schedule.SCHINSEXP, Schedule.SchTitltoFA, Schedule.SchTitletoClient, S"
    "chedule.LienHolder, Schedule.LienHolderNm, MstrLease.MLLessor, MstrLease.Client,"
    " Clients.ClientShNm, \"Edit\", Schedule.SchAH, IIf(IsNull([Assignee_Bank]),Null,"
    "[Assignee_Bank]), IIf(IsNull([SumOfOrigEquipCost]),Null,[SumOfOrigEquipCost]), I"
    "If(IsNull([SgrpPlb]),No,[SgrpPlb]), SchGrp.Assigned_Indicator, Schedule.SchFAID\015"
    "\012ORDER BY MstrLease.MLNo, Schedule.SCHNo1 DESC , Schedule.SchID DESC , Schedu"
    "le.SchNo DESC;\015\012"
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
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchFSL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchFAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Edit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bassign"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ECost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfUnitEstDelDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PLB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1004"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchTitleAssignee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SCHINSEXP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.LienHolder"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.LienHolderNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchMLRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchIn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchSigner2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchSigner3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchAssDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchTitleToTrust"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SCHINREC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SCHNo1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchAlias"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.Schsigned"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchTitltoFA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchTitletoClient"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchSigner"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchShortDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SCHINSAMT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLLessor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.Assigned_Indicator"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.Client"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchAH"
        dbLong "AggregateType" ="-1"
    End
End
