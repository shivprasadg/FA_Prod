Operation =1
Option =0
Where ="(((UnitTitleLog.EntryTypeId)=2))"
Having ="(((UnitTitleLog.EntryReasonId)<>7))"
Begin InputTables
    Name ="UnitTitleLog"
End
Begin OutputColumns
    Alias ="LastLog"
    Expression ="Max(UnitTitleLog.UnitTitleLogDate)"
    Expression ="UnitTitleLog.UnitRef"
    Expression ="UnitTitleLog.EntryTypeId"
    Expression ="UnitTitleLog.EntryReasonId"
    Expression ="UnitTitleLog.TitleNumber"
    Expression ="UnitTitleLog.UnitTitleZipcodeId"
    Expression ="UnitTitleLog.TitleSt"
    Alias ="ReasonOut"
    Expression ="Nz([Reason],\"\")"
    Expression ="UnitTitleLog.UnitTitleLogDate"
    Expression ="UnitTitleLog.ExpectedReturn"
End
Begin Groups
    Expression ="UnitTitleLog.UnitRef"
    GroupLevel =0
    Expression ="UnitTitleLog.EntryTypeId"
    GroupLevel =0
    Expression ="UnitTitleLog.EntryReasonId"
    GroupLevel =0
    Expression ="UnitTitleLog.TitleNumber"
    GroupLevel =0
    Expression ="UnitTitleLog.UnitTitleZipcodeId"
    GroupLevel =0
    Expression ="UnitTitleLog.TitleSt"
    GroupLevel =0
    Expression ="Nz([Reason],\"\")"
    GroupLevel =0
    Expression ="UnitTitleLog.UnitTitleLogDate"
    GroupLevel =0
    Expression ="UnitTitleLog.ExpectedReturn"
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
        dbText "Name" ="UnitTitleLog.UnitTitleLogDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="LastLog"
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
    Begin
        dbText "Name" ="UnitTitleLog.TitleNumber"
        dbInteger "ColumnWidth" ="1650"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleLog.UnitTitleZipcodeId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleLog.ExpectedReturn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleLog.EntryReasonId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleLog.Reason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleLog.TitleSt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ReasonOut"
        dbInteger "ColumnWidth" ="2955"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =65
    Top =44
    Right =1068
    Bottom =919
    Left =-1
    Top =-1
    Right =1414
    Bottom =574
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =45
        Top =25
        Right =402
        Bottom =361
        Top =0
        Name ="UnitTitleLog"
        Name =""
    End
End
