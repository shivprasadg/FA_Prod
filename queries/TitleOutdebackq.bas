Operation =1
Option =0
Having ="(((Last(UnitTitleLog.EntryTypeId))=2) AND ((Last(Reasons.ReasonId))=6 Or (Last(R"
    "easons.ReasonId))=8 Or (Last(Reasons.ReasonId))=9))"
Begin InputTables
    Name ="UnitTitleLog"
    Name ="Reasons"
End
Begin OutputColumns
    Expression ="UnitTitleLog.UnitRef"
    Alias ="LastOfUnitTitleLogDate"
    Expression ="Last(UnitTitleLog.UnitTitleLogDate)"
    Alias ="LastOfEntryTypeId"
    Expression ="Last(UnitTitleLog.EntryTypeId)"
    Alias ="LastOfReasonId"
    Expression ="Last(Reasons.ReasonId)"
    Expression ="Reasons.ReasonDescription"
End
Begin Joins
    LeftTable ="UnitTitleLog"
    RightTable ="Reasons"
    Expression ="UnitTitleLog.EntryReasonId = Reasons.ReasonId"
    Flag =1
End
Begin Groups
    Expression ="UnitTitleLog.UnitRef"
    GroupLevel =0
    Expression ="Reasons.ReasonDescription"
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
        dbText "Name" ="UnitTitleLog.UnitRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastOfUnitTitleLogDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastOfEntryTypeId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Reasons.ReasonDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastOfReasonId"
        dbInteger "ColumnWidth" ="4230"
        dbBoolean "ColumnHidden" ="0"
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
    Bottom =277
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =803
        Bottom =156
        Top =0
        Name ="UnitTitleLog"
        Name =""
    End
    Begin
        Left =851
        Top =12
        Right =995
        Bottom =156
        Top =0
        Name ="Reasons"
        Name =""
    End
End
