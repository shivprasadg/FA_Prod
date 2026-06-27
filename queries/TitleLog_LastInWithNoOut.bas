Operation =1
Option =0
Where ="(((UnitTitleHistoryLastRecv.UnitTitleLogDate)>[UnitTitleHistoryLastSent].[UnitTi"
    "tleLogDate]))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="UnitTitleHistoryLastRecv"
    Name ="UnitTitleHistoryLastSent"
End
Begin OutputColumns
    Expression ="vw_SixKeys.ClientShortName"
    Expression ="vw_SixKeys.Schedule"
    Expression ="vw_SixKeys.UnitGroup"
    Expression ="vw_SixKeys.Status"
    Expression ="vw_SixKeys.AssetID"
    Expression ="vw_SixKeys.VIN"
    Alias ="Direction"
    Expression ="\"IN\""
    Expression ="UnitTitleHistoryLastRecv.TitleNumber"
    Expression ="UnitTitleHistoryLastRecv.EntryReason"
    Expression ="UnitTitleHistoryLastRecv.UnitTitleLogDate"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="UnitTitleHistoryLastRecv"
    Expression ="vw_SixKeys.AssetID = UnitTitleHistoryLastRecv.UnitRef"
    Flag =1
    LeftTable ="UnitTitleHistoryLastRecv"
    RightTable ="UnitTitleHistoryLastSent"
    Expression ="UnitTitleHistoryLastRecv.UnitRef = UnitTitleHistoryLastSent.UnitRef"
    Flag =2
End
Begin OrderBy
    Expression ="vw_SixKeys.ClientShortName"
    Flag =0
    Expression ="vw_SixKeys.Schedule"
    Flag =0
    Expression ="vw_SixKeys.UnitGroup"
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
        dbText "Name" ="Direction"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastTitleEntry_In.UnitTitleLogDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastTitleEntry_In.Action"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastTitleEntry_In.EntryReason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleHistoryLastEntry.EntryTypeId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastTitleEntry_In.TitleNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastTitleEntry_Out.Action"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastTitleEntry_In.UnitRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastTitleEntry_Out.UnitTitleLogDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastTitleEntry_Out.EntryReason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastTitleEntry_Out.UnitRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastTitleEntry_In.UnitTitleLogID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.ClientShortName"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleHistoryLastEntry.TitleNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleHistoryLastEntry.EntryReason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleHistoryLastEntry.UnitTitleLogDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleHistoryLastRecv.EntryReason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleHistoryLastEntry_1.EntryTypeId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleHistoryLastRecv.TitleNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleHistoryLastRecv.UnitTitleLogDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleHistoryLastSent.UnitRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleHistoryLastRecv.UnitTitleLogID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleHistoryLastSent.UnitTitleLogID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =87
    Top =27
    Right =1615
    Bottom =707
    Left =-1
    Top =-1
    Right =1510
    Bottom =397
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =40
        Top =21
        Right =184
        Bottom =357
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =285
        Top =100
        Right =429
        Bottom =350
        Top =0
        Name ="UnitTitleHistoryLastRecv"
        Name =""
    End
    Begin
        Left =509
        Top =103
        Right =653
        Bottom =329
        Top =0
        Name ="UnitTitleHistoryLastSent"
        Name =""
    End
End
