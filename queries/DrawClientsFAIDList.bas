dbMemo "SQL" ="SELECT Clients.clientID, Clients.ClientShNm AS Client, dbo_vw_Sch_Grp_Units_Chil"
    "dParts.FAID, IIf(Count([drawdowns].[FAID])>0,\"Yes\",\"No\") AS Drawn\015\012FRO"
    "M (Clients INNER JOIN dbo_vw_Sch_Grp_Units_ChildParts ON (Clients.clientID = dbo"
    "_vw_Sch_Grp_Units_ChildParts.clientID) AND (Clients.clientID = dbo_vw_Sch_Grp_Un"
    "its_ChildParts.clientID)) LEFT JOIN DrawDowns ON dbo_vw_Sch_Grp_Units_ChildParts"
    ".FAID = DrawDowns.FAID\015\012GROUP BY Clients.clientID, Clients.ClientShNm, dbo"
    "_vw_Sch_Grp_Units_ChildParts.FAID, Clients.clientType\015\012HAVING (((Clients.c"
    "lientType)=\"CLIENT\"))\015\012ORDER BY Clients.clientID DESC , Clients.ClientSh"
    "Nm;\015\012"
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
    End
    Begin
        dbText "Name" ="Drawn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_vw_Sch_Grp_Units_ChildParts.FAID"
        dbLong "AggregateType" ="-1"
    End
End
