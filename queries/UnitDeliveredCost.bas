Operation =1
Option =0
Begin InputTables
    Name ="Units"
    Name ="ChildParts"
End
Begin OutputColumns
    Alias ="GroupId"
    Expression ="Units.GroupREF"
    Alias ="AssetId"
    Expression ="Units.UnitID"
    Alias ="SyndRent"
    Expression ="Units.UnitRent"
    Alias ="LORent"
    Expression ="Units.UnitSRent"
    Alias ="UnitCost"
    Expression ="Units.unitfinalcost"
    Alias ="ChildCost"
    Expression ="Sum(CCur(Nz([childFinalCost],0)))"
    Alias ="DeliveredCost"
    Expression ="CCur([unitfinalcost]+Sum(Nz([childFinalCost],0)))"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="ChildParts"
    Expression ="Units.UnitID = ChildParts.UnitREF"
    Flag =2
End
Begin Groups
    Expression ="Units.GroupREF"
    GroupLevel =0
    Expression ="Units.UnitID"
    GroupLevel =0
    Expression ="Units.UnitRent"
    GroupLevel =0
    Expression ="Units.UnitSRent"
    GroupLevel =0
    Expression ="Units.unitfinalcost"
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
        dbText "Name" ="GrouipId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DeliveredCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AssetId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LORent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SyndRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupId"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =43
    Top =91
    Right =1645
    Bottom =974
    Left =-1
    Top =-1
    Right =1570
    Bottom =515
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =124
        Top =72
        Right =268
        Bottom =216
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =355
        Top =122
        Right =596
        Bottom =594
        Top =0
        Name ="ChildParts"
        Name =""
    End
End
