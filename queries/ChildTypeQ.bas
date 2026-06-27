Operation =1
Option =0
Where ="(((lookup_groupTypes.grpClass)=\"other\"))"
Begin InputTables
    Name ="ChildParts"
    Name ="lookup_groupTypes"
End
Begin OutputColumns
    Expression ="ChildParts.UnitREF"
    Expression ="ChildParts.ChildMake"
    Expression ="lookup_groupTypes.grpClass"
End
Begin Joins
    LeftTable ="ChildParts"
    RightTable ="lookup_groupTypes"
    Expression ="ChildParts.ChildType = lookup_groupTypes.groupType"
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
        dbText "Name" ="ChildParts.UnitREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lookup_groupTypes.grpClass"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1035
    Bottom =604
    Left =-1
    Top =-1
    Right =1003
    Bottom =222
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="ChildParts"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="lookup_groupTypes"
        Name =""
    End
End
