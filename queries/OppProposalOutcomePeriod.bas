dbMemo "SQL" ="SELECT Clients.LOinitial, Clients.ClientShNm AS Client, Deals.FAID, DatePart('q'"
    ",[dealOutcomeDate]) AS DealQTR, DatePart('yyyy',[dealOutcomeDate]) AS DealYear, "
    "Deals.dealOutcome, Deals.dealLPAFSent AS LPAFSent, Deals.dealOutcomeDate, Round("
    "[deals].[dealGrandTotal]) AS dealGrandTotal, Deals.dealShortDesc, Deals.dealoutc"
    "ome AS RptGrp, Deals.DealLOEx, lookup_transTypes.TransType, IIf([DealLOEx]=\"hs\""
    " Or [DealLOEx]=\"sm\",\"HS\",[DealLOEx]) AS RptLOEx, Deals.dealUnitQty, Deals.de"
    "alSchedNum, Clients.clienttype, \"x\" AS Grpon, IIf([dealoutcome]=\"awarded\",1,"
    "IIf([dealoutcome]=\"lost\",2,3)) AS Outcomsrt, DatePart(\"yyyy\",Date()) AS yr, "
    "Clients.ClientGroupID\015\012FROM (Clients INNER JOIN Salesperson ON Clients.LOi"
    "nitial = Salesperson.SPInitials) INNER JOIN (Deals INNER JOIN lookup_transTypes "
    "ON Deals.transTypeRef = lookup_transTypes.TransTypeID) ON Clients.clientID = Dea"
    "ls.clientRef\015\012WHERE (((Clients.ClientGroupID)<> 18))\015\012ORDER BY Clien"
    "ts.LOinitial, Clients.ClientShNm, Deals.dealOutcome;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Clients.LOinitial"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="RptLOEx"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="yr"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="20"
    End
    Begin
        dbText "Name" ="Deals.FAID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="Deals.dealShortDesc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="RptGrp"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="Deals.dealUnitQty"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="15"
    End
    Begin
        dbText "Name" ="Deals.dealSchedNum"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="16"
    End
    Begin
        dbText "Name" ="Clients.clienttype"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="17"
    End
    Begin
        dbText "Name" ="Grpon"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="18"
    End
    Begin
        dbText "Name" ="Deals.dealOutcome"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="Deals.dealOutcomeDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="lookup_transTypes.TransType"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="Outcomsrt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="19"
    End
    Begin
        dbText "Name" ="DealYear"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="Deals.DealLOEx"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="DealQTR"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="LPAFSent"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dealGrandTotal"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
End
