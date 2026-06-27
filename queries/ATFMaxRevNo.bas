Operation =1
Option =0
Begin InputTables
    Name ="ATFRevT"
End
Begin OutputColumns
    Alias ="SchID"
    Expression ="ATFRevT.schref"
    Alias ="RevNo"
    Expression ="Max(ATFRevT.RevNo)"
    Expression ="ATFRevT.RevDt"
End
Begin Groups
    Expression ="ATFRevT.schref"
    GroupLevel =0
    Expression ="ATFRevT.RevDt"
    GroupLevel =0
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
        dbText "Name" ="SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RevNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFRevT.RevDt"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1238
    Bottom =1038
    Left =-1
    Top =-1
    Right =1206
    Bottom =681
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =57
        Top =79
        Right =201
        Bottom =345
        Top =0
        Name ="ATFRevT"
        Name =""
    End
End
