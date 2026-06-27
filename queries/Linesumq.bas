dbMemo "SQL" ="SELECT Lines.Linebank, Lines.Lineid, Lines.LineDt, Lines.LineType, Lines.LineMin"
    "Loan, [lineamt]+[lineinc1]+[lineinc2]+[lineinc3]+[lineinc4]+[lineinc5] AS lineto"
    "t, IIf(IsNull([SumOfPNAmt]),0,[SumOfPNAmt]) AS Use\015\012FROM Lines LEFT JOIN P"
    "ntSumq ON Lines.Lineid=PntSumq.Lineid\015\012WHERE (((Lines.Linebank)=[forms]![b"
    "ank_view]![clientid]));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Lines.Linebank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="linetot"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Use"
        dbLong "AggregateType" ="-1"
    End
End
