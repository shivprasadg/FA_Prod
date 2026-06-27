Operation =1
Option =0
Where ="(((FirstTitleLog.EntryTypeId)=1))"
Begin InputTables
    Name ="FirstTitleLog"
End
Begin OutputColumns
    Expression ="FirstTitleLog.UnitTitleLogId"
    Expression ="FirstTitleLog.UnitRef"
    Expression ="FirstTitleLog.EntryTypeId"
    Expression ="FirstTitleLog.EntryReasonId"
    Expression ="FirstTitleLog.TitleNumber"
    Expression ="FirstTitleLog.UnitTitleZipcodeId"
    Expression ="FirstTitleLog.TitleSt"
    Expression ="FirstTitleLog.ReasonLogged"
    Expression ="FirstTitleLog.UnitTitleLogDate"
    Expression ="FirstTitleLog.ExpectedReturn"
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
        dbText "Name" ="FirstTitleLog.ExpectedReturn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstTitleLog.TitleNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstTitleLog.UnitTitleZipcodeId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstTitleLog.EntryReasonId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstTitleLog.EntryTypeId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstTitleLog.TitleSt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstTitleLog.UnitRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstTitleLog.UnitTitleLogId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstTitleLog.ReasonLogged"
        dbInteger "ColumnWidth" ="1980"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstTitleLog.UnitTitleLogDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1430
    Bottom =974
    Left =-1
    Top =-1
    Right =1398
    Bottom =650
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =70
        Top =11
        Right =519
        Bottom =365
        Top =0
        Name ="FirstTitleLog"
        Name =""
    End
End
