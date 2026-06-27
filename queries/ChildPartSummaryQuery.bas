Operation =1
Option =0
Having ="(((ChildParts.ExclFrGrpOEC)=False))"
Begin InputTables
    Name ="SchGrp"
    Name ="Units"
    Name ="ChildParts"
End
Begin OutputColumns
    Expression ="SchGrp.SGrpID"
    Expression ="ChildParts.UnitREF"
    Alias ="SumOfchildFinalCost"
    Expression ="Sum(ChildParts.childFinalCost)"
    Expression ="ChildParts.ExclFrGrpOEC"
    Alias ="SumOfChildDeposit"
    Expression ="Sum(ChildParts.ChildDeposit)"
End
Begin Joins
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
    LeftTable ="Units"
    RightTable ="ChildParts"
    Expression ="Units.UnitID = ChildParts.UnitREF"
    Flag =1
End
Begin Groups
    Expression ="SchGrp.SGrpID"
    GroupLevel =0
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
        dbText "Name" ="SumOfchildFinalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ExclFrGrpOEC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfChildDeposit"
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
    Bottom =266
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =240
        Top =264
        Right =384
        Bottom =408
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =48
        Top =264
        Right =192
        Bottom =408
        Top =0
        Name ="Units"
        Name =""
    End
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
