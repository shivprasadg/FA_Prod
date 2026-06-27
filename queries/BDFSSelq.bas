dbMemo "SQL" ="SELECT Emps.empint AS SalespersonID, Emps.empnm AS SPInitials, Emps.[LO Exec], E"
    "mps.FS\015\012FROM Emps\015\012WHERE (((Emps.[LO Exec])=True)) OR (((Emps.FS)=Tr"
    "ue))\015\012ORDER BY Emps.empnm;\015\012"
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
        dbText "Name" ="SalespersonID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SPInitials"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Emps.[LO Exec]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Emps.FS"
        dbLong "AggregateType" ="-1"
    End
End
