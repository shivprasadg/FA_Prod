Operation =1
Option =0
Begin InputTables
    Name ="LookUp_ChildPartMake"
    Name ="LookUp_ChildPartType"
    Name ="LookUp_ChildPartModel"
End
Begin OutputColumns
    Alias ="MasterTypeID"
    Expression ="LookUp_ChildPartType.ChildPartTypeID"
    Expression ="LookUp_ChildPartType.ChildPartType"
    Expression ="LookUp_ChildPartMake.ChildPartTypeID"
    Expression ="LookUp_ChildPartMake.ChildPartMake"
    Expression ="LookUp_ChildPartModel.ChildPartMakeID"
    Expression ="LookUp_ChildPartModel.ChildPartModel"
    Expression ="LookUp_ChildPartModel.ChildPartModelID"
End
Begin Joins
    LeftTable ="LookUp_ChildPartMake"
    RightTable ="LookUp_ChildPartType"
    Expression ="LookUp_ChildPartMake.ChildPartTypeID = LookUp_ChildPartType.ChildPartTypeID"
    Flag =1
    LeftTable ="LookUp_ChildPartMake"
    RightTable ="LookUp_ChildPartModel"
    Expression ="LookUp_ChildPartMake.ChildPartMakeID = LookUp_ChildPartModel.ChildPartMakeID"
    Flag =2
End
Begin OrderBy
    Expression ="LookUp_ChildPartType.ChildPartType"
    Flag =0
    Expression ="LookUp_ChildPartMake.ChildPartMake"
    Flag =0
    Expression ="LookUp_ChildPartModel.ChildPartModel"
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
        dbText "Name" ="LookUp_ChildPartType.ChildPartType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LookUp_ChildPartMake.ChildPartTypeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LookUp_ChildPartModel.ChildPartMakeID"
        dbInteger "ColumnWidth" ="1980"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LookUp_ChildPartMake.ChildPartMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LookUp_ChildPartModel.ChildPartModel"
        dbInteger "ColumnWidth" ="2220"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MasterTypeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LookUp_ChildPartModel.ChildPartModelID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-198
    Top =99
    Right =1279
    Bottom =1077
    Left =-1
    Top =-1
    Right =1459
    Bottom =701
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =255
        Top =200
        Right =399
        Bottom =344
        Top =0
        Name ="LookUp_ChildPartMake"
        Name =""
    End
    Begin
        Left =28
        Top =115
        Right =172
        Bottom =259
        Top =0
        Name ="LookUp_ChildPartType"
        Name =""
    End
    Begin
        Left =497
        Top =270
        Right =641
        Bottom =414
        Top =0
        Name ="LookUp_ChildPartModel"
        Name =""
    End
End
