Operation =1
Option =0
Where ="(((LastTitleLogInOut.entrytypeid)=2))"
Begin InputTables
    Name ="UnitTitleLog"
    Name ="LookUp_TitleReasons"
    Name ="LastTitleLogInOut"
End
Begin OutputColumns
    Alias ="Action"
    Expression ="\"Out\""
    Expression ="UnitTitleLog.UnitRef"
    Expression ="UnitTitleLog.EntryReason"
    Expression ="UnitTitleLog.TitleNumber"
    Expression ="UnitTitleLog.TitleSt"
    Expression ="UnitTitleLog.UnitTitleLogDate"
    Expression ="UnitTitleLog.ExpectedReturn"
    Expression ="UnitTitleLog.UserID"
    Expression ="UnitTitleLog.LastUpdate"
    Expression ="UnitTitleLog.UnitTitleLogID"
    Expression ="UnitTitleLog.EntryTypeId"
    Expression ="UnitTitleLog.EntryReasonId"
End
Begin Joins
    LeftTable ="UnitTitleLog"
    RightTable ="LookUp_TitleReasons"
    Expression ="UnitTitleLog.EntryReasonId = LookUp_TitleReasons.ReasonId"
    Flag =1
    LeftTable ="UnitTitleLog"
    RightTable ="LastTitleLogInOut"
    Expression ="UnitTitleLog.UnitTitleLogID = LastTitleLogInOut.UnitTitleLogID"
    Flag =1
End
Begin OrderBy
    Expression ="UnitTitleLog.UnitRef"
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
        dbText "Name" ="UnitTitleLog.UnitRef"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1125"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Action"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1020"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="UnitTitleLog.TitleSt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1020"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="UnitTitleLog.UnitTitleLogID"
        dbInteger "ColumnWidth" ="1710"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleLog.UnitTitleLogDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleLog.ExpectedReturn"
        dbInteger "ColumnWidth" ="1905"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleLog.EntryTypeId"
        dbInteger "ColumnWidth" ="1530"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleLog.EntryReasonId"
        dbInteger "ColumnWidth" ="1755"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleLog.EntryReason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleLog.TitleNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleLog.UserID"
        dbInteger "ColumnWidth" ="1035"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleLog.LastUpdate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =12
    Top =23
    Right =1270
    Bottom =669
    Left =-1
    Top =-1
    Right =1240
    Bottom =239
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =286
        Top =12
        Right =485
        Bottom =208
        Top =0
        Name ="UnitTitleLog"
        Name =""
    End
    Begin
        Left =722
        Top =36
        Right =866
        Bottom =180
        Top =0
        Name ="LookUp_TitleReasons"
        Name =""
    End
    Begin
        Left =35
        Top =0
        Right =179
        Bottom =144
        Top =0
        Name ="LastTitleLogInOut"
        Name =""
    End
End
