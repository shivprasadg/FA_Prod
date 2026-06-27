Operation =1
Option =0
Where ="(((lookup_groupTypes.grpClass)=\"body\"))"
Begin InputTables
    Name ="ChildParts"
    Name ="lookup_groupTypes"
End
Begin OutputColumns
    Expression ="ChildParts.UnitREF"
    Expression ="ChildParts.ChildVIN"
    Expression ="ChildParts.ChildYear"
    Expression ="ChildParts.ChildMake"
    Expression ="ChildParts.ChildClass"
    Expression ="ChildParts.ChildType"
    Expression ="ChildParts.EstimatedDeliveryDate"
    Expression ="lookup_groupTypes.grpClass"
    Expression ="ChildParts.ActualDeliveryDate"
End
Begin Joins
    LeftTable ="ChildParts"
    RightTable ="lookup_groupTypes"
    Expression ="ChildParts.ChildType = lookup_groupTypes.groupType"
    Flag =1
End
Begin OrderBy
    Expression ="ChildParts.UnitREF"
    Flag =0
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
        dbText "Name" ="ChildParts.ChildYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lookup_groupTypes.grpClass"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.EstimatedDeliveryDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ActualDeliveryDate"
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
    Bottom =398
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =501
        Bottom =408
        Top =0
        Name ="ChildParts"
        Name =""
    End
    Begin
        Left =576
        Top =18
        Right =720
        Bottom =162
        Top =0
        Name ="lookup_groupTypes"
        Name =""
    End
End
