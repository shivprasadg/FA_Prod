Operation =1
Option =0
Where ="(((LastTitleEntry_Out.EntryReason) Is Null) AND ((LastTitleEntry_Out.UnitTitleLo"
    "gDate) Is Null))"
Begin InputTables
    Name ="LastTitleEntry_In"
    Name ="LastTitleEntry_Out"
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Expression ="vw_SixKeys.ClientShortName"
    Expression ="vw_SixKeys.Schedule"
    Expression ="vw_SixKeys.UnitGroup"
    Expression ="vw_SixKeys.Status"
    Expression ="LastTitleEntry_In.UnitTitleLogID"
    Expression ="LastTitleEntry_In.UnitRef"
    Expression ="LastTitleEntry_In.Action"
    Expression ="LastTitleEntry_In.EntryReason"
    Expression ="LastTitleEntry_In.TitleNumber"
    Expression ="LastTitleEntry_In.UnitTitleLogDate"
End
Begin Joins
    LeftTable ="LastTitleEntry_In"
    RightTable ="LastTitleEntry_Out"
    Expression ="LastTitleEntry_In.UnitRef = LastTitleEntry_Out.UnitRef"
    Flag =2
    LeftTable ="vw_SixKeys"
    RightTable ="LastTitleEntry_In"
    Expression ="vw_SixKeys.AssetID = LastTitleEntry_In.UnitRef"
    Flag =1
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
        dbText "Name" ="LastTitleEntry_In.TitleNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastTitleEntry_In.UnitRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastTitleEntry_In.UnitTitleLogID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.ClientShortName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="LastTitleEntry_Out.Action"
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
End
Begin
    State =0
    Left =74
    Top =25
    Right =1602
    Bottom =705
    Left =-1
    Top =-1
    Right =1510
    Bottom =414
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =193
        Top =82
        Right =337
        Bottom =442
        Top =0
        Name ="LastTitleEntry_In"
        Name =""
    End
    Begin
        Left =390
        Top =102
        Right =534
        Bottom =356
        Top =0
        Name ="LastTitleEntry_Out"
        Name =""
    End
    Begin
        Left =40
        Top =21
        Right =184
        Bottom =357
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
End
