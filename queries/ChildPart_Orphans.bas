Operation =1
Option =0
Where ="(((Units.UnitID) Is Null))"
Begin InputTables
    Name ="ChildParts"
    Name ="Units"
End
Begin OutputColumns
    Expression ="ChildParts.ChildPartID"
    Expression ="ChildParts.ChildYear"
    Expression ="ChildParts.ChildMakeRef"
    Expression ="ChildParts.ChildMake"
    Expression ="ChildParts.ChildType"
    Expression ="ChildParts.ChildStatus"
    Expression ="ChildParts.ChildVIN"
    Expression ="ChildParts.childFinalCost"
    Expression ="Units.UnitID"
    Expression ="ChildParts.UnitREF"
End
Begin Joins
    LeftTable ="ChildParts"
    RightTable ="Units"
    Expression ="ChildParts.UnitREF = Units.UnitID"
    Flag =2
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
        dbText "Name" ="ChildParts.ChildPartID"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildYear"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildMakeRef"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1725"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="ChildParts.ChildMake"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildType"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildStatus"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildVIN"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.childFinalCost"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.UnitREF"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =109
    Top =3
    Right =1719
    Bottom =717
    Left =-1
    Top =-1
    Right =1578
    Bottom =606
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =165
        Top =85
        Right =353
        Bottom =612
        Top =0
        Name ="ChildParts"
        Name =""
    End
    Begin
        Left =440
        Top =93
        Right =584
        Bottom =237
        Top =0
        Name ="Units"
        Name =""
    End
End
