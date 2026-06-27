dbMemo "SQL" ="SELECT UnitsAndChildUnionQuery.unitID, Sum(UnitsAndChildUnionQuery.UnitFinalCost"
    ") AS SumOfUnitFinalCost, Groups.groupID, Groups.dealRef\015\012FROM UnitsAndChil"
    "dUnionQuery INNER JOIN (UnitsSummary INNER JOIN Groups ON UnitsSummary.GroupREF="
    "Groups.groupID) ON UnitsAndChildUnionQuery.unitID=UnitsSummary.UnitID\015\012GRO"
    "UP BY UnitsAndChildUnionQuery.unitID, Groups.groupID, Groups.dealRef;\015\012"
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
        dbText "Name" ="Groups.groupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.dealRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsAndChildUnionQuery.unitID"
        dbLong "AggregateType" ="-1"
    End
End
