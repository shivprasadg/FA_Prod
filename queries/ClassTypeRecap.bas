Operation =1
Option =9
Begin InputTables
    Name ="CLASS"
    Name ="lookup_groupTypes"
End
Begin OutputColumns
    Expression ="lookup_groupTypes.grpClass"
    Expression ="lookup_groupTypes.groupType"
    Expression ="lookup_groupTypes.GrpTypAbrev"
    Expression ="CLASS.TITLED"
End
Begin Joins
    LeftTable ="CLASS"
    RightTable ="lookup_groupTypes"
    Expression ="CLASS.CLASS = lookup_groupTypes.grpClass"
    Flag =1
End
Begin OrderBy
    Expression ="lookup_groupTypes.grpClass"
    Flag =0
    Expression ="lookup_groupTypes.groupType"
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
        dbText "Name" ="lookup_groupTypes.grpClass"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lookup_groupTypes.groupType"
        dbInteger "ColumnWidth" ="2295"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lookup_groupTypes.GrpTypAbrev"
        dbInteger "ColumnWidth" ="1635"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lookup_groupTypes.groupTypeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lookup_groupTypes.upsize_ts"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CLASS.TITLED"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =114
    Top =178
    Right =891
    Bottom =776
    Left =-1
    Top =-1
    Right =745
    Bottom =213
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="CLASS"
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
