Operation =1
Option =0
Having ="(((ChildParts.ExclFrGrpOEC)=False))"
Begin InputTables
    Name ="ChildParts"
End
Begin OutputColumns
    Expression ="ChildParts.UnitREF"
    Alias ="ChildPartCost"
    Expression ="Sum(ChildParts.childFinalCost)"
    Alias ="ChildDeposit"
    Expression ="Sum(ChildParts.ChildDeposit)"
End
Begin OrderBy
    Expression ="ChildParts.UnitREF"
    Flag =0
End
Begin Groups
    Expression ="ChildParts.UnitREF"
    GroupLevel =0
    Expression ="ChildParts.ExclFrGrpOEC"
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
        dbText "Name" ="ChildPartCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildDeposit"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =38
    Top =86
    Right =1290
    Bottom =776
    Left =-1
    Top =-1
    Right =1220
    Bottom =424
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =418
        Top =26
        Right =682
        Bottom =267
        Top =0
        Name ="ChildParts"
        Name =""
    End
End
