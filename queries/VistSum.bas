dbMemo "SQL" ="SELECT TempSC.clientID, CFOVISTQ.CFOVisti, Treasvisits.TreasVisit, FinAnal.AnalV"
    "isit, Trans.TransMgrVisit, FleetMgr.FMgrvisit, ShopMgrq.ShopMgrVisit\015\012FROM"
    " (((((CFOVISTQ RIGHT JOIN TempSC ON CFOVISTQ.clientID=TempSC.clientID) LEFT JOIN"
    " Treasvisits ON TempSC.clientID=Treasvisits.clientID) LEFT JOIN Trans ON TempSC."
    "clientID=Trans.clientID) LEFT JOIN FleetMgr ON TempSC.clientID=FleetMgr.clientID"
    ") LEFT JOIN FinAnal ON TempSC.clientID=FinAnal.clientID) LEFT JOIN ShopMgrq ON T"
    "empSC.clientID=ShopMgrq.clientID;\015\012"
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
        dbText "Name" ="CFOVISTQ.CFOVisti"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Treasvisits.TreasVisit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FinAnal.AnalVisit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Trans.TransMgrVisit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetMgr.FMgrvisit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ShopMgrq.ShopMgrVisit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TempSC.clientID"
        dbLong "AggregateType" ="-1"
    End
End
