dbMemo "SQL" ="SELECT UnitDescByGroup_A.GroupREF AS GroupID, Count(UnitDescByGroup_A.GroupREF) "
    "AS GroupCount, Max(UnitDescByGroup_A.Desc) AS ADesc, Min(UnitDescByGroup_A.Desc)"
    " AS BDesc, IIf([GroupCount]>1,\"(\" & [AUnits] & \") \" & [ADesc] & \" (\" & [BU"
    "nits] & \") \" & [BDesc],\"(\" & [AUnits] & \") \" & [ADesc]) AS GroupDesc, Max("
    "UnitDescByGroup_A.Units) AS AUnits, Min(UnitDescByGroup_A.Units) AS BUnits INTO "
    "UnitDescByGroupTemp\015\012FROM UnitDescByGroup_A\015\012GROUP BY UnitDescByGrou"
    "p_A.GroupREF;\015\012"
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
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="GroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AUnits"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BUnits"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ADesc"
        dbInteger "ColumnWidth" ="3705"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BDesc"
        dbInteger "ColumnWidth" ="3705"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupDesc"
        dbInteger "ColumnWidth" ="7365"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupCount"
        dbLong "AggregateType" ="-1"
    End
End
