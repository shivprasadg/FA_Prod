dbMemo "SQL" ="SELECT Schedule.SchID, GroupFinalizedTracker.GroupID, SchGrp.SGrpNO AS GrpNo, [S"
    "GrpNO] & \" - \" & Format([RequestCompleted],\"mm/dd/yy\") AS TMTSigned\015\012F"
    "ROM Schedule LEFT JOIN (SchGrp LEFT JOIN GroupFinalizedTracker ON SchGrp.SGrpID "
    "= GroupFinalizedTracker.GroupID) ON Schedule.SchID = SchGrp.SGrpSchID\015\012WHE"
    "RE (((GroupFinalizedTracker.Department)=\"TMT\" Or (GroupFinalizedTracker.Depart"
    "ment) Is Null))\015\012GROUP BY Schedule.SchID, GroupFinalizedTracker.GroupID, S"
    "chGrp.SGrpNO, [SGrpNO] & \" - \" & Format([RequestCompleted],\"mm/dd/yy\");\015\012"
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
        dbText "Name" ="GroupFinalizedTracker.Department"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFinalizedTracker.GroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TMTSigned"
        dbInteger "ColumnWidth" ="2655"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GrpNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
    End
End
