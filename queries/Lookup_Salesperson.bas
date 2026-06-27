Operation =1
Option =0
Where ="(((Emps.[LO Exec])=True))"
Begin InputTables
    Name ="Emps"
End
Begin OutputColumns
    Expression ="Emps.EmpInt"
    Expression ="Emps.EmpNm"
End
Begin OrderBy
    Expression ="Emps.EmpInt"
    Flag =0
End
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
        dbText "Name" ="Emps.EmpInt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Emps.EmpNm"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1258
    Bottom =686
    Left =-1
    Top =-1
    Right =1226
    Bottom =204
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =316
        Top =12
        Right =460
        Bottom =156
        Top =0
        Name ="Emps"
        Name =""
    End
End
