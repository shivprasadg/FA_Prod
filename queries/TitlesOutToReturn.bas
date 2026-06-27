Operation =1
Option =0
Where ="(((LastTitleLog.EntryTypeId)=2) AND ((IIf(Not IsNull([UnitTitleRecd]) And [Entry"
    "ReasonId]=18,False,[ExpectedReturn]))=True))"
Begin InputTables
    Name ="LastTitleLog"
    Name ="Units"
End
Begin OutputColumns
    Expression ="LastTitleLog.UnitTitleLogId"
    Expression ="LastTitleLog.UnitRef"
    Expression ="LastTitleLog.EntryTypeId"
    Expression ="LastTitleLog.EntryReasonId"
    Expression ="LastTitleLog.TitleNumber"
    Expression ="LastTitleLog.TitleSt"
    Expression ="LastTitleLog.ReasonLogged"
    Expression ="LastTitleLog.UnitTitleLogDate"
    Alias ="ToReturn"
    Expression ="IIf(Not IsNull([UnitTitleRecd]) And [EntryReasonId]=18,False,[ExpectedReturn])"
End
Begin Joins
    LeftTable ="LastTitleLog"
    RightTable ="Units"
    Expression ="LastTitleLog.UnitRef = Units.UnitID"
    Flag =1
End
Begin OrderBy
    Expression ="LastTitleLog.UnitTitleLogDate"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="LastTitleLog.UnitTitleLogId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastTitleLog.TitleNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastTitleLog.UnitTitleLogDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastTitleLog.UnitRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastTitleLog.EntryTypeId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastTitleLog.TitleSt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastTitleLog.EntryReasonId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastTitleLog.ReasonLogged"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4380"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="ToReturn"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =101
    Top =140
    Right =1514
    Bottom =764
    Left =-1
    Top =-1
    Right =1381
    Bottom =582
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =256
        Bottom =251
        Top =0
        Name ="LastTitleLog"
        Name =""
    End
    Begin
        Left =304
        Top =12
        Right =448
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
End
