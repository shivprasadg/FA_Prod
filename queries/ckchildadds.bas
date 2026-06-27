Operation =1
Option =0
Where ="(((ChildParts.ChildMake)=\"vomela\") AND ((ChildParts.childFinalCost)=351.37) AN"
    "D ((Units.GroupREF)=1417))"
Begin InputTables
    Name ="ChildParts"
    Name ="Units"
End
Begin OutputColumns
    Expression ="ChildParts.UnitREF"
    Expression ="ChildParts.ChildPartID"
    Expression ="ChildParts.ChildMake"
    Expression ="ChildParts.childFinalCost"
End
Begin Joins
    LeftTable ="ChildParts"
    RightTable ="Units"
    Expression ="ChildParts.UnitREF = Units.UnitID"
    Flag =1
End
Begin OrderBy
    Expression ="ChildParts.UnitREF"
    Flag =0
    Expression ="ChildParts.ChildPartID"
    Flag =1
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
        dbText "Name" ="ChildParts.ChildMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.childFinalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildPartID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.UnitREF"
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
    Bottom =436
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =334
        Top =40
        Right =478
        Bottom =184
        Top =0
        Name ="ChildParts"
        Name =""
    End
    Begin
        Left =132
        Top =14
        Right =276
        Bottom =158
        Top =0
        Name ="Units"
        Name =""
    End
End
