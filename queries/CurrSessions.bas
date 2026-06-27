dbMemo "SQL" ="SELECT qry_login.User_name, qry_login.MaxOfTime_In_Stamp AS Time_In, IIf([MaxOfT"
    "ime_OUT_Stamp]<[MaxOfTime_In_Stamp],\"\",[MaxOfTime_OUT_Stamp]) AS Time_Out, Fir"
    "st(TBL_Employees.[Active User]) AS [FirstOfActive User]\015\012FROM (qry_login L"
    "EFT JOIN qry_logout ON qry_login.User_name=qry_logout.User_name) LEFT JOIN TBL_E"
    "mployees ON qry_logout.User_name=TBL_Employees.[User ID]\015\012GROUP BY qry_log"
    "in.User_name, qry_login.MaxOfTime_In_Stamp, IIf([MaxOfTime_OUT_Stamp]<[MaxOfTime"
    "_In_Stamp],\"\",[MaxOfTime_OUT_Stamp]);\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbInteger "RowHeight" ="285"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbMemo "OrderBy" ="[Currsessions].[Time_Out] DESC"
Begin
    Begin
        dbText "Name" ="Time_Out"
        dbInteger "ColumnWidth" ="2655"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Time_In"
        dbInteger "ColumnWidth" ="3225"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_login.User_name"
        dbInteger "ColumnWidth" ="2280"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfActive User"
        dbLong "AggregateType" ="-1"
    End
End
