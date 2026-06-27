dbMemo "SQL" ="SELECT \"\" & Chr(13) & Chr(10) & IIf([dtype]=\"DIR\",\"All\",[G].[SGrpNO]) AS s"
    "GrpNo, DocLinks.*\015\012FROM DocLinks LEFT JOIN (SELECT sGrpNo, SGrpId FROM Sch"
    "Grp)  AS G ON DocLinks.GroupID = G.SGrpID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
End
