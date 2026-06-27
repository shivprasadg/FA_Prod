dbMemo "SQL" ="SELECT DISTINCT Bills.BillID, MstrLease.MLNo, Schedule.SchNo, Schedule.SchID, Sc"
    "hGrp.SGrpID, IIf(IsNull([BLCD]),[Forms]![Billing]![BillingSubForm].[Form]![billt"
    "o],IIf([Forms]![Billing]![BillingSubForm].[Form]![billto]>[BLCD],[BLCD],[Forms]!"
    "[Billing]![BillingSubForm].[Form]![billto])) AS BillEndDt, Bills.Billfrom AS Bil"
    "lstart, IIf([billto]>IIf(IsNull([BLCD]),[schcrexp],IIf([schcrexp]>[BLCD],[BLCD],"
    "[schcrexp])),IIf(IsNull([BLCD]),[schcrexp],IIf([schcrexp]>[BLCD],[BLCD],[schcrex"
    "p])),[billto]) AS SlvEnd, IIf(IsNull([BLCD]),[billto],IIf([billto]>[BLCD],[BLCD]"
    ",[billto])) AS PDEND\015\012FROM Bills INNER JOIN ((MstrLease INNER JOIN Schedul"
    "e ON MstrLease.MLID = Schedule.SchMLRef) INNER JOIN SchGrp ON Schedule.SchID = S"
    "chGrp.SGrpSchID) ON Bills.BillSchRef = Schedule.SchID\015\012WHERE (((Bills.Bill"
    "ID)=[Forms]![Billing]![BillingSubForm].[Form]![BillID]));\015\012"
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
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillEndDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Billstart"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SlvEnd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2715"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDEnd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bills.BillID"
        dbLong "AggregateType" ="-1"
    End
End
