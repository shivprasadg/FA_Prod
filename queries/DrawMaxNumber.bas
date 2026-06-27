Operation =1
Option =0
Begin InputTables
    Name ="DrawDowns"
End
Begin OutputColumns
    Expression ="DrawDowns.DrawRequestNumber"
    Alias ="DrawNumber"
    Expression ="Max(DrawDowns.DrawNumber)"
End
Begin Groups
    Expression ="DrawDowns.DrawRequestNumber"
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
        dbText "Name" ="DrawNumber"
        dbInteger "ColumnWidth" ="3825"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawDowns.DrawRequestNumber"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1617
    Bottom =991
    Left =-1
    Top =-1
    Right =1585
    Bottom =469
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =89
        Top =36
        Right =383
        Bottom =342
        Top =0
        Name ="DrawDowns"
        Name =""
    End
End
