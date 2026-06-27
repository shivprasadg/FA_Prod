Operation =1
Option =0
Begin InputTables
    Name ="UnitTitleLog"
    Name ="Reasons"
End
Begin OutputColumns
    Alias ="TitleLogId"
    Expression ="Last(UnitTitleLog.UnitTitleLogId)"
    Alias ="AssetId"
    Expression ="UnitTitleLog.UnitRef"
    Alias ="LogDate"
    Expression ="Max(UnitTitleLog.UnitTitleLogDate)"
    Alias ="Action"
    Expression ="Last(IIf([entrytypeid]=1,\"In\",\"Out\"))"
    Alias ="Description"
    Expression ="Last(Reasons.ReasonDescription)"
    Alias ="TitleNumber"
    Expression ="Last(UnitTitleLog.TitleNumber)"
    Alias ="ExpectedReturn"
    Expression ="Last(UnitTitleLog.ExpectedReturn)"
End
Begin Joins
    LeftTable ="UnitTitleLog"
    RightTable ="Reasons"
    Expression ="UnitTitleLog.EntryReasonId = Reasons.ReasonId"
    Flag =2
End
Begin OrderBy
    Expression ="UnitTitleLog.UnitRef"
    Flag =1
End
Begin Groups
    Expression ="UnitTitleLog.UnitRef"
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
        dbText "Name" ="Action"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Description"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3660"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="LogDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TitleLogId"
        dbInteger "ColumnWidth" ="2310"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TitleNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExpectedReturn"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =94
    Top =183
    Right =1352
    Bottom =775
    Left =-1
    Top =-1
    Right =1226
    Bottom =252
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =485
        Bottom =156
        Top =0
        Name ="UnitTitleLog"
        Name =""
    End
    Begin
        Left =533
        Top =12
        Right =677
        Bottom =156
        Top =0
        Name ="Reasons"
        Name =""
    End
End
