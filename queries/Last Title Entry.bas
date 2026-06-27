Operation =1
Option =0
Begin InputTables
    Name ="UnitTitleLog"
    Name ="LookUp_TitleReasons"
End
Begin OutputColumns
    Alias ="AssetID"
    Expression ="UnitTitleLog.UnitRef"
    Expression ="UnitTitleLog.UnitRef"
    Alias ="TitleLogId"
    Expression ="Last(UnitTitleLog.UnitTitleLogId)"
    Alias ="LogDate"
    Expression ="Max(UnitTitleLog.UnitTitleLogDate)"
    Alias ="Action"
    Expression ="Last(IIf([entrytypeid]=1,\"In\",\"Out\"))"
    Alias ="LastOfTitleNumber"
    Expression ="Last(UnitTitleLog.TitleNumber)"
    Alias ="LastOfUnitTitleLogDate"
    Expression ="Max(UnitTitleLog.UnitTitleLogDate)"
    Expression ="LookUp_TitleReasons.ReasonDescription"
End
Begin Joins
    LeftTable ="UnitTitleLog"
    RightTable ="LookUp_TitleReasons"
    Expression ="UnitTitleLog.EntryReasonId = LookUp_TitleReasons.ReasonId"
    Flag =2
End
Begin OrderBy
    Expression ="UnitTitleLog.UnitRef"
    Flag =0
End
Begin Groups
    Expression ="UnitTitleLog.UnitRef"
    GroupLevel =0
    Expression ="UnitTitleLog.UnitRef"
    GroupLevel =0
    Expression ="LookUp_TitleReasons.ReasonDescription"
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
        dbText "Name" ="LastOfTitleNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastOfUnitTitleLogDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleLog.UnitRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LookUp_TitleReasons.ReasonDescription"
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
    Bottom =156
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
        Left =749
        Top =35
        Right =893
        Bottom =179
        Top =0
        Name ="LookUp_TitleReasons"
        Name =""
    End
End
