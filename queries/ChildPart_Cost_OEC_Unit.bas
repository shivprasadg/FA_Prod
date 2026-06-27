Operation =1
Option =0
Begin InputTables
    Name ="ChildParts"
End
Begin OutputColumns
    Alias ="UnitId"
    Expression ="ChildParts.UnitREF"
    Alias ="ChildPartCost"
    Expression ="CCur(Nz([childFinalCost],0))"
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
        dbText "Name" ="UnitId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildPartCost"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =43
    Top =91
    Right =1309
    Bottom =974
    Left =-1
    Top =-1
    Right =1234
    Bottom =464
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =285
        Top =138
        Right =579
        Bottom =355
        Top =0
        Name ="ChildParts"
        Name =""
    End
End
