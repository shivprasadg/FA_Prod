dbMemo "SQL" ="SELECT Clients.clientID, Clients.ClientShNm AS Client, IIf(Count([drawdowns].[Dr"
    "awRequestNumber])>0,\"Yes\",\"No\") AS Draws\015\012FROM (Clients INNER JOIN vw_"
    "Sch_Grp_Units_ChildParts ON (Clients.clientID = vw_Sch_Grp_Units_ChildParts.clie"
    "ntID) AND (Clients.clientID = vw_Sch_Grp_Units_ChildParts.clientID)) LEFT JOIN D"
    "rawDowns ON vw_Sch_Grp_Units_ChildParts.MLID = DrawDowns.MLID\015\012GROUP BY Cl"
    "ients.clientID, Clients.ClientShNm, Clients.clientType\015\012HAVING (((Clients."
    "clientType)=\"CLIENT\"))\015\012ORDER BY Clients.ClientShNm;\015\012"
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
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2115"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Draws"
        dbLong "AggregateType" ="-1"
    End
End
