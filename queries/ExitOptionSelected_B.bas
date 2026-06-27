Operation =1
Option =0
Begin InputTables
    Name ="ExitOptionSelected_A"
End
Begin OutputColumns
    Expression ="ExitOptionSelected_A.UnitID"
    Alias ="Exc"
    Expression ="Max(Nz([Exch],\"----\"))"
    Alias ="Extd"
    Expression ="Max(Nz([Extend],\"----\"))"
    Alias ="End"
    Expression ="Max(Nz([Term],\"----\"))"
End
Begin Groups
    Expression ="ExitOptionSelected_A.UnitID"
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
        dbText "Name" ="ExitOptionSelected_A.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Exc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Extd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="End"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =113
    Top =116
    Right =1441
    Bottom =997
    Left =-1
    Top =-1
    Right =1296
    Bottom =563
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =70
        Top =63
        Right =214
        Bottom =207
        Top =0
        Name ="ExitOptionSelected_A"
        Name =""
    End
End
