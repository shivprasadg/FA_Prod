Operation =1
Option =0
Begin InputTables
    Name ="Units"
    Name ="ChildParts"
End
Begin OutputColumns
    Alias ="GroupID"
    Expression ="Units.GroupREF"
    Alias ="UnitCount"
    Expression ="Count(Units.UnitID)"
    Alias ="UnitCost"
    Expression ="Sum(Nz([UnitFinalCost],0))"
    Alias ="ChildPartCost"
    Expression ="Sum(Nz([childFinalCost],0))"
    Alias ="SyndRent"
    Expression ="Sum(Nz([UNITRENT],0))"
    Alias ="LORent"
    Expression ="Sum(Nz([UNITSRENT],0))"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="ChildParts"
    Expression ="Units.UnitID = ChildParts.UnitREF"
    Flag =1
End
Begin Groups
    Expression ="Units.GroupREF"
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
        dbText "Name" ="GroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SyndRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LORent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.childFinalCost"
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
    Right =1326
    Bottom =974
    Left =-1
    Top =-1
    Right =1251
    Bottom =566
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =260
        Top =98
        Right =571
        Bottom =375
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =626
        Top =174
        Right =882
        Bottom =529
        Top =0
        Name ="ChildParts"
        Name =""
    End
End
