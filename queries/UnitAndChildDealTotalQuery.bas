dbMemo "SQL" ="SELECT Groups.dealRef, Sum(UnitsAndChildUnionQuery.UnitFinalCost) AS SumOfUnitFi"
    "nalCost\015\012FROM UnitsAndChildUnionQuery INNER JOIN (UnitsSummary INNER JOIN "
    "Groups ON UnitsSummary.GroupREF=Groups.groupID) ON UnitsAndChildUnionQuery.unitI"
    "D=UnitsSummary.UnitID\015\012GROUP BY Groups.dealRef;\015\012"
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
        dbText "Name" ="SumOfUnitFinalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.dealRef"
        dbLong "AggregateType" ="-1"
    End
End
