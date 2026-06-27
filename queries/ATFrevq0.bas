Operation =1
Option =0
Where ="(((ATFRevT.RevNo) Is Not Null) AND ((ATFRevT.ReviewDt) Is Null))"
Begin InputTables
    Name ="ATFRevT"
End
Begin OutputColumns
    Expression ="ATFRevT.schref"
    Expression ="ATFRevT.RevNo"
    Expression ="ATFRevT.RevDt"
    Expression ="ATFRevT.ReviewDt"
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
        dbText "Name" ="ATFRevT.schref"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFRevT.RevNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFRevT.RevDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFRevT.ReviewDt"
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
    Bottom =237
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="ATFRevT"
        Name =""
    End
End
