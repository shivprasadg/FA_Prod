dbMemo "SQL" ="SELECT Salesperson.SPInitials AS SM, Clients.clientCompanyName AS [Co Name], Cli"
    "ents.clientCity AS City, Clients.clientState AS ST, Clients.clientZip AS Zip, Cl"
    "ients.fleetSize, Clients.fleetDaycabs AS [No of DC], Clients.fleetDaycabsMPY AS "
    "[DC MPY], Clients.fleetSleepers AS [No of Slp], Clients.fleetSleepersMPY AS SlpM"
    "PY, Clients.FleetReplacementCycle AS [Repl Cycle], Clients.creditType, Clients.c"
    "reditMoodyRating AS Moodys, Clients.creditRank AS [FA Rank], [SGRD]+[DCGRD]+[DCM"
    "IGRD]+[SPLMIGRD]+[MAINTGRD]+[LCGRD]+[OBCGRD]+[MODGRD]+IIf([creditRank]=\"tba\",["
    "MOODYGRD],[FACRGRD]) AS nscore, Clients.clienttype\015\012FROM (Salesperson INNE"
    "R JOIN Clients ON Salesperson.SalespersonID=Clients.SalespersonREF) INNER JOIN N"
    "SCOREQ ON Clients.clientID=NSCOREQ.clientID\015\012WHERE (((Clients.clienttype)="
    "\"target\"))\015\012ORDER BY Clients.clientCompanyName;\015\012"
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
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="Clients.creditType"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3015"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="sT"
        dbText "Format" =">"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="sm"
        dbText "Format" =">"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
        dbInteger "ColumnWidth" ="705"
        dbBoolean "ColumnHidden" ="0"
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
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="DC MPY"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="No of Slp"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="SlpMPY"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="Repl Cycle"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="Moodys"
        dbInteger "ColumnWidth" ="2205"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="FA Rank"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="15"
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
        dbLong "AggregateType" ="-1"
    End
End
