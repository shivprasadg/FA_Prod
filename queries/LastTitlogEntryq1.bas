Operation =1
Option =0
Begin InputTables
    Name ="Units"
    Name ="UnitTitleLog"
End
Begin OutputColumns
    Expression ="UnitTitleLog.UnitRef"
    Alias ="MaxOfUnitTitleLogDate"
    Expression ="Max(UnitTitleLog.UnitTitleLogDate)"
    Alias ="MaxOfUnitTitleLogId"
    Expression ="Max(UnitTitleLog.UnitTitleLogId)"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="UnitTitleLog"
    Expression ="Units.UnitID = UnitTitleLog.UnitRef"
    Flag =1
End
Begin OrderBy
    Expression ="UnitTitleLog.UnitRef"
    Flag =0
End
Begin Groups
    Expression ="UnitTitleLog.UnitRef"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbMemo "OrderBy" ="[LastTitlogEntryq1].[UnitRef]"
Begin
    Begin
        dbText "Name" ="UnitTitleLog.UnitRef"
        dbInteger "ColumnWidth" ="1125"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="MaxOfUnitTitleLogDate"
        dbInteger "ColumnWidth" ="2595"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="MaxOfUnitTitleLogId"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4260"
        dbBoolean "ColumnHidden" ="0"
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
    Bottom =362
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =417
        Top =53
        Right =561
        Bottom =184
        Top =0
        Name ="UnitTitleLog"
        Name =""
    End
End
