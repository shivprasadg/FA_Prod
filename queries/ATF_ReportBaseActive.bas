dbMemo "SQL" ="SELECT DLookUp(\"ATFID\",\"Schedule\",\"Schid=\" & [vw_SixKeys].[Schid]) AS ATFI"
    "D, vw_SixKeys.FAID, vw_SixKeys.MLID, vw_SixKeys.ClientId AS LesseeID, 1 AS Credi"
    "tRank, DLookUp(\"ATFSchSpecID\",\"Schedule\",\"Schid=\" & [vw_SixKeys].[schid]) "
    "AS ATFSchSpecID, vw_SixKeys.BankShortName AS Assignee_Bank, Min(\"FAID: \" & [vw"
    "_SixKeys].[FAID] & \"-\" & [dealEnteredDate]) AS ATFDesc, vw_SixKeys.CompanyName"
    " AS Lessee, Count(vw_SixKeys.UnitID) AS ActualUnits, Min(Deals.dealEnteredDate) "
    "AS ATFDate, Min(Deals.dealLPAFcomp) AS LPAFDate, Deals.DealLOEx AS BDEX, vw_SixK"
    "eys.MLNo, DLookUp(\"clientTaxID\",\"Clients\",\"ClientID=\" & [vw_SixKeys].[Clie"
    "ntID]) AS ClientTaxID, 0 AS CoLesseeID, vw_SixKeys.BankId, ECount(\"SchId\",\"vw"
    "_SixKeys\",\"BankId= \" & [vw_SixKeys].[BankId] & \" AND FAID='\" & [vw_Sixkeys]"
    ".[FAID] & \"'\",True) AS NumOfSchedules, ECount(\"SGrpId\",\"vw_SixKeys\",\"Bank"
    "Id= \" & [vw_SixKeys].[BankId] & \" AND FAID='\" & [vw_Sixkeys].[FAID] & \"'\",T"
    "rue) AS NumOfGroups, vw_SixKeys.SchID, Deals.dealOutcome\015\012FROM vw_SixKeys "
    "LEFT JOIN Deals ON vw_SixKeys.FAID = Deals.FAID\015\012WHERE (((vw_SixKeys.MLOri"
    "g)=\"FA\"))\015\012GROUP BY DLookUp(\"ATFID\",\"Schedule\",\"Schid=\" & [vw_SixK"
    "eys].[Schid]), vw_SixKeys.FAID, vw_SixKeys.MLID, vw_SixKeys.ClientId, 1, DLookUp"
    "(\"ATFSchSpecID\",\"Schedule\",\"Schid=\" & [vw_SixKeys].[schid]), vw_SixKeys.Ba"
    "nkShortName, vw_SixKeys.CompanyName, Deals.DealLOEx, vw_SixKeys.MLNo, DLookUp(\""
    "clientTaxID\",\"Clients\",\"ClientID=\" & [vw_SixKeys].[ClientID]), 0, vw_SixKey"
    "s.BankId, ECount(\"SchId\",\"vw_SixKeys\",\"BankId= \" & [vw_SixKeys].[BankId] &"
    " \" AND FAID='\" & [vw_Sixkeys].[FAID] & \"'\",True), ECount(\"SGrpId\",\"vw_Six"
    "Keys\",\"BankId= \" & [vw_SixKeys].[BankId] & \" AND FAID='\" & [vw_Sixkeys].[FA"
    "ID] & \"'\",True), vw_SixKeys.SchID, Deals.dealOutcome;\015\012"
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
        dbText "Name" ="ATFDesc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3465"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Lessee"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2520"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="NumOfGroups"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NumOfSchedules"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2040"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="LPAFDate"
        dbInteger "ColumnWidth" ="2085"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="BDEX"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ActualUnits"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.BankId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CreditRank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LesseeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFSchSpecID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Assignee_Bank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CoLesseeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClientTaxID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealOutcome"
        dbLong "AggregateType" ="-1"
    End
End
