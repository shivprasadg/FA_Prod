Operation =1
Option =0
Begin InputTables
    Name ="ChildParts"
End
Begin OutputColumns
    Expression ="ChildParts.UnitREF"
    Alias ="SumOfchildFinalCost"
    Expression ="Sum(ChildParts.childFinalCost)"
    Expression ="ChildParts.ChildType"
End
Begin Groups
    Expression ="ChildParts.UnitREF"
    GroupLevel =0
    Expression ="ChildParts.ChildType"
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
        dbText "Name" ="ChildParts.UnitREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfchildFinalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildType"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =76
    Top =132
    Right =1252
    Bottom =772
    Left =-1
    Top =-1
    Right =1144
    Bottom =451
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =312
        Bottom =253
        Top =0
        Name ="ChildParts"
        Name =""
    End
End
