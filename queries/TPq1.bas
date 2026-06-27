dbMemo "SQL" ="SELECT IIf([CLIENTTYPE]=\"TARGET\",1,IIf([CLIENTTYPE]=\"2ND TARGET\",2,3)) AS SR"
    "TORDER, Clients.clienttype, Clients.loinitial AS SM, Clients.clientCompanyName A"
    "S [Co Name], Clients.clientCity AS City, Clients.clientState AS ST, Clients.clie"
    "ntZip AS Zip, Clients.fleetSize, Clients.fleetDaycabs AS [No of DC], NSCOREQ.DCM"
    "PY, NSCOREQ.DCMIO, Clients.fleetSleepers AS [No of Slp], NSCOREQ.SlpMPY, NSCOREQ"
    ".SLPMIO, Clients.FleetReplacementCycle AS [Repl Cycle], Clients.creditType, Clie"
    "nts.creditMoodyRating AS Moodys, Clients.creditRank AS [FA Rank], [SGRD]+[DCMIGR"
    "D]+[SPLMIGRD]+[MAINTGRD]+[LCGRD]+[OBCGRD]+[MODGRD]+IIf([creditRank]=\"tba\",[MOO"
    "DYGRD],[FACRGRD]) AS nscore, Clients.creditFinance, Clients.creditMaint, Clients"
    ".clientobc, Clients.fleetDaycabsMPY, Clients.fleetSleepersMPY, [clients].[fleetd"
    "aycabs]+[clients].[fleetsleepers] AS FS, GROUPRTQ.maxofcreditfascore AS GRPMAX, "
    "Clients.ClientGroupID, Clients.JFApproval, IIf([clients].[fleetdaycabs]+[clients"
    "].[fleetsleepers]=0 Or ([clients].[fleetdaycabs]>0 And [clients].[fleetdaycabsmp"
    "y]=0) Or ([clients].[fleetsleepers]>0 And [clients].[fleetsleepersmpy]=0) Or [cl"
    "ientobc]=\"tba\" Or [creditfinance]=\"tbd\" Or [creditmaint]=\"tbd\" Or IsNull(["
    "fleetreplacementcycle]) Or ([creditrank]=\"tb\" And [creditmoodyrating]=\"tba\" "
    "And IsNull([jfapproval])),True,0) AS Mdata, TQ2.ex, TQ2.DT, Clients.creditRank, "
    "Clients.creditFAProfile, Clients.creditFAScore, Clients.creditFAScore\015\012FRO"
    "M (((Clients LEFT JOIN NSCOREQ ON Clients.clientID=NSCOREQ.clientID) LEFT JOIN G"
    "ROUPRTQ ON Clients.ClientGroupID=GROUPRTQ.ClientGroupID) LEFT JOIN TQ2 ON Client"
    "s.clientID=TQ2.clientID) INNER JOIN Salesperson ON Clients.LOinitial=Salesperson"
    ".SPInitials\015\012GROUP BY IIf([CLIENTTYPE]=\"TARGET\",1,IIf([CLIENTTYPE]=\"2ND"
    " TARGET\",2,3)), Clients.clienttype, Clients.loinitial, Clients.clientCompanyNam"
    "e, Clients.clientCity, Clients.clientState, Clients.clientZip, Clients.fleetSize"
    ", Clients.fleetDaycabs, NSCOREQ.DCMPY, NSCOREQ.DCMIO, Clients.fleetSleepers, NSC"
    "OREQ.SlpMPY, NSCOREQ.SLPMIO, Clients.FleetReplacementCycle, Clients.creditType, "
    "Clients.creditMoodyRating, Clients.creditRank, [SGRD]+[DCMIGRD]+[SPLMIGRD]+[MAIN"
    "TGRD]+[LCGRD]+[OBCGRD]+[MODGRD]+IIf([creditRank]=\"tba\",[MOODYGRD],[FACRGRD]), "
    "Clients.creditFinance, Clients.creditMaint, Clients.clientobc, Clients.fleetDayc"
    "absMPY, Clients.fleetSleepersMPY, [clients].[fleetdaycabs]+[clients].[fleetsleep"
    "ers], GROUPRTQ.maxofcreditfascore, Clients.ClientGroupID, Clients.JFApproval, II"
    "f([clients].[fleetdaycabs]+[clients].[fleetsleepers]=0 Or ([clients].[fleetdayca"
    "bs]>0 And [clients].[fleetdaycabsmpy]=0) Or ([clients].[fleetsleepers]>0 And [cl"
    "ients].[fleetsleepersmpy]=0) Or [clientobc]=\"tba\" Or [creditfinance]=\"tbd\" O"
    "r [creditmaint]=\"tbd\" Or IsNull([fleetreplacementcycle]) Or ([creditrank]=\"tb"
    "\" And [creditmoodyrating]=\"tba\" And IsNull([jfapproval])),True,0), TQ2.ex, TQ"
    "2.DT, Clients.creditRank, Clients.creditFAProfile, Clients.creditFAScore, Client"
    "s.creditFAScore, Clients.ClientGroupID\015\012HAVING (((Clients.clienttype) Like"
    " \"*target\") AND ((Clients.ClientGroupID)<> 18))\015\012ORDER BY Clients.loinit"
    "ial, Clients.clientCompanyName;\015\012"
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
        dbText "Name" ="Clients.fleetSize"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.creditType"
        dbInteger "ColumnWidth" ="3015"
        dbInteger "ColumnOrder" ="15"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sT"
        dbInteger "ColumnOrder" ="5"
        dbText "Format" =">"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sm"
        dbInteger "ColumnWidth" ="705"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" =">"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Co Name"
        dbInteger "ColumnWidth" ="3795"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="City"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="No of DC"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="No of Slp"
        dbInteger "ColumnOrder" ="12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Repl Cycle"
        dbInteger "ColumnOrder" ="14"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Moodys"
        dbInteger "ColumnWidth" ="2205"
        dbInteger "ColumnOrder" ="16"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FA Rank"
        dbInteger "ColumnOrder" ="17"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clienttype"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Zip"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="nscore"
        dbInteger "ColumnOrder" ="20"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NSCOREQ.DCMPY"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NSCOREQ.DCMIO"
        dbInteger "ColumnOrder" ="18"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NSCOREQ.SlpMPY"
        dbInteger "ColumnOrder" ="13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NSCOREQ.SLPMIO"
        dbInteger "ColumnOrder" ="19"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.creditFinance"
        dbInteger "ColumnWidth" ="1275"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.creditMaint"
        dbInteger "ColumnWidth" ="3675"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientobc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.fleetDaycabsMPY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.fleetSleepersMPY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.[CLIENT GROUP]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GRPMAX"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.JFApproval"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SRTORDER"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mdata"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TQ2.ex"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TQ2.DT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.creditRank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.creditFAProfile"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.creditFAScore"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1033"
        dbLong "AggregateType" ="-1"
    End
End
