Operation =1
Option =0
Where ="(((UnitTitleHistoryLastEntry.ExpectedReturn)=True))"
Begin InputTables
    Name ="UnitTitleHistoryLastEntry"
    Name ="vw_SixKeys"
    Name ="Units"
End
Begin OutputColumns
    Alias ="Client"
    Expression ="vw_SixKeys.ClientShortName"
    Expression ="vw_SixKeys.BDEX"
    Alias ="Assignee"
    Expression ="vw_SixKeys.BankShortName"
    Expression ="vw_SixKeys.Schedule"
    Expression ="vw_SixKeys.UnitGroup"
    Expression ="vw_SixKeys.AssetID"
    Alias ="UnitNum"
    Expression ="vw_SixKeys.Unitnum"
    Expression ="vw_SixKeys.VIN"
    Expression ="vw_SixKeys.Status"
    Alias ="Date Logged"
    Expression ="UnitTitleHistoryLastEntry.UnitTitleLogDate"
    Alias ="Reason Out"
    Expression ="UnitTitleHistoryLastEntry.EntryReason"
    Alias ="Title Number"
    Expression ="UnitTitleHistoryLastEntry.TitleNumber"
    Alias ="Reg State"
    Expression ="UnitTitleHistoryLastEntry.TitledLocationState"
    Alias ="DaysOut"
    Expression ="DateDiff(\"d\",[UnitTitleLogDate],Date())"
    Expression ="Units.LienRelease"
End
Begin Joins
    LeftTable ="UnitTitleHistoryLastEntry"
    RightTable ="vw_SixKeys"
    Expression ="UnitTitleHistoryLastEntry.UnitRef = vw_SixKeys.AssetID"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="Units"
    Expression ="vw_SixKeys.UnitID = Units.UnitID"
    Flag =1
End
Begin OrderBy
    Expression ="vw_SixKeys.ClientShortName"
    Flag =0
    Expression ="vw_SixKeys.Schedule"
    Flag =0
    Expression ="vw_SixKeys.UnitGroup"
    Flag =0
    Expression ="UnitTitleHistoryLastEntry.UnitTitleLogDate"
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
        dbText "Name" ="vw_SixKeys.VIN"
        dbInteger "ColumnWidth" ="2805"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitGroup"
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
        dbText "Name" ="Date Logged"
        dbInteger "ColumnWidth" ="0"
        dbBoolean "ColumnHidden" ="-1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Title Number"
        dbInteger "ColumnWidth" ="2220"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Reason Out"
        dbInteger "ColumnWidth" ="2070"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Assignee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.BDEX"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Reg State"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Reg State]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DaysOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.LienRelease"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-947
    Top =237
    Right =980
    Bottom =1029
    Left =-1
    Top =-1
    Right =1903
    Bottom =153
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =283
        Bottom =156
        Top =0
        Name ="UnitTitleHistoryLastEntry"
        Name =""
    End
    Begin
        Left =438
        Top =0
        Right =582
        Bottom =264
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =630
        Top =12
        Right =774
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
End
