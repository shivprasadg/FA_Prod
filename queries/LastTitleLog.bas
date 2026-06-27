Operation =1
Option =0
Begin InputTables
    Name ="UnitTitleLog"
    Alias ="TL"
    Name ="SELECT Max(UnitTitleLog.UnitTitleLogDate) AS MaxLogDate, UnitREF FROM UnitTitleL"
        "og GROUP BY UnitRef"
    Alias ="T2"
End
Begin OutputColumns
    Expression ="TL.UnitTitleLogId"
    Expression ="TL.UnitRef"
    Expression ="TL.EntryTypeId"
    Expression ="TL.EntryReasonId"
    Expression ="TL.TitleNumber"
    Expression ="TL.TitleSt"
    Alias ="ReasonLogged"
    Expression ="Nz([EntryReason],'')"
    Expression ="TL.UnitTitleLogDate"
    Expression ="TL.ExpectedReturn"
End
Begin Joins
    LeftTable ="TL"
    RightTable ="T2"
    Expression ="TL.UnitRef = T2.UnitREF"
    Flag =1
    LeftTable ="TL"
    RightTable ="T2"
    Expression ="TL.UnitTitleLogDate = T2.MaxLogDate"
    Flag =1
End
Begin OrderBy
    Expression ="TL.UnitTitleLogDate"
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
        dbText "Name" ="TL.UnitTitleLogDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TL.EntryReasonId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TL.TitleNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TL.UnitTitleLogId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TL.UnitRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TL.EntryTypeId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TL.TitleSt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TL.ExpectedReturn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ReasonLogged"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4335"
        dbBoolean "ColumnHidden" ="0"
    End
End
Begin
    State =0
    Left =5
    Top =223
    Right =1435
    Bottom =1032
    Left =-1
    Top =-1
    Right =1398
    Bottom =287
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =414
        Top =77
        Right =558
        Bottom =221
        Top =0
        Name ="TL"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="T2"
        Name =""
    End
End
