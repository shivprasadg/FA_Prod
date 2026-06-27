Operation =1
Option =0
Begin InputTables
    Name ="UnitTitleLog"
End
Begin OutputColumns
    Expression ="UnitTitleLog.UnitRef"
    Alias ="UnitTitleLogID"
    Expression ="Max(UnitTitleLog.UnitTitleLogID)"
    Expression ="UnitTitleLog.EntryTypeId"
End
Begin OrderBy
    Expression ="UnitTitleLog.UnitRef"
    Flag =0
End
Begin Groups
    Expression ="UnitTitleLog.UnitRef"
    GroupLevel =0
    Expression ="UnitTitleLog.EntryTypeId"
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
        dbText "Name" ="UnitTitleLogID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleLog.UnitRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleLog.EntryTypeId"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =136
    Top =0
    Right =1483
    Bottom =863
    Left =-1
    Top =-1
    Right =1329
    Bottom =586
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =77
        Top =19
        Right =221
        Bottom =308
        Top =0
        Name ="UnitTitleLog"
        Name =""
    End
End
