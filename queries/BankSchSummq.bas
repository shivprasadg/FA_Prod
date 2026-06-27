dbMemo "SQL" ="SELECT MstrLease.MLOrig, MstrLease.MLID, MstrLease.MLNo, Schedule.SchID, Schedul"
    "e.SchFSL, Schedule.SchNo, SchGrp.SGrpNO, Schedule.SchOut, Schedule.SchIn, Schedu"
    "le.SchAssDt, Schedule.SchShortDesc, Schedule.SchFAID, Sum(SchGrp.OrigEquipCost) "
    "AS SumOfOrigEquipCost, MstrLease.Client, Clients.ClientShNm, SchGrp.Assigned_Ind"
    "icator, SchGrp.Assignee_Bank, SchGrp.BLCD, SchGrp.LeasePymtForGroup, SchGrp.Leas"
    "eType, SchGrp.UnitsInGroup, SchGrp.OrigEquipCost, SchGrp.LeaseTermPrimary, SchGr"
    "p.LeasePymtForGroup, SchGrp.AssignmentDate, DateAdd(\"m\",[LeaseTermPrimary],[BL"
    "CD])-1 AS Exp, SchGrp.ExchangeITGroup, SchGrp.ExchangeITMonth, IIf([ExchangeITGr"
    "oup]=False,\"none\",\"@ Mo \" & [ExchangeITMonth]) AS ExIT, SchGrp.Assignee_Bank"
    ", SchGrp.AssignmentDate\015\012FROM ((Clients INNER JOIN MstrLease ON Clients.cl"
    "ientID = MstrLease.Client) INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchM"
    "LRef) INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID\015\012GROUP BY Mst"
    "rLease.MLOrig, MstrLease.MLID, MstrLease.MLNo, Schedule.SchID, Schedule.SchFSL, "
    "Schedule.SchNo, SchGrp.SGrpNO, Schedule.SchOut, Schedule.SchIn, Schedule.SchAssD"
    "t, Schedule.SchShortDesc, Schedule.SchFAID, MstrLease.Client, Clients.ClientShNm"
    ", SchGrp.Assigned_Indicator, SchGrp.BLCD, SchGrp.LeaseType, SchGrp.UnitsInGroup,"
    " SchGrp.OrigEquipCost, SchGrp.LeaseTermPrimary, SchGrp.LeasePymtForGroup, SchGrp"
    ".ExchangeITGroup, SchGrp.ExchangeITMonth, IIf([ExchangeITGroup]=False,\"none\",\""
    "@ Mo \" & [ExchangeITMonth]), SchGrp.Assignee_Bank, SchGrp.AssignmentDate, SchGr"
    "p.LeasePymtForGroup, SchGrp.Assignee_Bank, SchGrp.AssignmentDate, SchGrp.LeasePy"
    "mtForGroup, MstrLease.MLOrig, SchGrp.AssignmentDate\015\012HAVING (((MstrLease.M"
    "LNo) Not Like \"*demo*\") AND ((SchGrp.Assignee_Bank) Is Not Null) AND ((MstrLea"
    "se.MLOrig)=\"fa\"))\015\012ORDER BY MstrLease.MLNo, Schedule.SchID DESC , Schedu"
    "le.SchNo DESC , SchGrp.SGrpNO;\015\012"
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
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="MstrLease.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1800"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="930"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Schedule.SchOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchIn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchAssDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchShortDesc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4125"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Schedule.SchFAID"
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
        dbText "Name" ="MstrLease.MLID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="885"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Schedule.SchFSL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1155"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Expr1018"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Exp"
        dbInteger "ColumnWidth" ="1155"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExIT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.Assignee_Bank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.Assigned_Indicator"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.UnitsInGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.AssignmentDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1016"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1024"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfOrigEquipCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.ExchangeITMonth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LeasePymtForGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.ExchangeITGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LeaseType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.OrigEquipCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LeaseTermPrimary"
        dbLong "AggregateType" ="-1"
    End
End
