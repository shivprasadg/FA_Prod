dbMemo "SQL" ="SELECT vw_SixKeys.ClientShortName, vw_SixKeys.Schedule, Sum(Units.UnitIFamt) AS "
    "SumOfUnitIFamt, Min(Units.unitifdate) AS MinOfunitifdate, DateAdd(\"m\",6,Min([U"
    "nitIFDate])) AS Expr1\015\012FROM (Units INNER JOIN vw_SixKeys ON Units.UnitID ="
    " vw_SixKeys.UnitID) INNER JOIN SchGrp ON vw_SixKeys.SGrpID = SchGrp.SGrpID\015\012"
    "GROUP BY vw_SixKeys.ClientShortName, vw_SixKeys.Schedule;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="SumOfUnitIFamt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitIFamt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.ClientShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MinOfunitifdate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitifdate"
        dbLong "AggregateType" ="-1"
    End
End
