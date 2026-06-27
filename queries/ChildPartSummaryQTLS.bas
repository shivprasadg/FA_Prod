Operation =1
Option =0
Having ="(((ChildParts.ExclFrGrpOEC)=No) AND ((Units.UnitMake)=\"peterbilt\"))"
Begin InputTables
    Name ="Units"
    Name ="ChildParts"
End
Begin OutputColumns
    Expression ="ChildParts.UnitREF"
    Alias ="SumOfchildFinalCost"
    Expression ="Sum(ChildParts.childFinalCost)"
    Expression ="ChildParts.ExclFrGrpOEC"
    Expression ="Units.UnitMake"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="ChildParts"
    Expression ="Units.UnitID = ChildParts.UnitREF"
    Flag =1
End
Begin Groups
    Expression ="ChildParts.UnitREF"
    GroupLevel =0
    Expression ="ChildParts.ExclFrGrpOEC"
    GroupLevel =0
    Expression ="Units.UnitMake"
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
        dbText "Name" ="ChildParts.ExclFrGrpOEC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitMake"
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
    Bottom =286
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =360
        Top =12
        Right =504
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
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
