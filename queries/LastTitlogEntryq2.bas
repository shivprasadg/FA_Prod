Operation =1
Option =0
Where ="(((UnitTitleLog.EntryTypeId)=2) AND ((UnitTitleLog.EntryReasonId)<>4 And (UnitTi"
    "tleLog.EntryReasonId)<>7))"
Begin InputTables
    Name ="Units"
    Name ="LastTitlogEntryq1"
    Name ="UnitTitleLog"
End
Begin OutputColumns
    Expression ="LastTitlogEntryq1.UnitRef"
    Expression ="UnitTitleLog.EntryTypeId"
    Expression ="UnitTitleLog.EntryReasonId"
    Expression ="LastTitlogEntryq1.MaxOfUnitTitleLogId"
    Expression ="Units.UNITVIN"
End
Begin Joins
    LeftTable ="LastTitlogEntryq1"
    RightTable ="UnitTitleLog"
    Expression ="LastTitlogEntryq1.MaxOfUnitTitleLogDate = UnitTitleLog.UnitTitleLogDate"
    Flag =1
    LeftTable ="LastTitlogEntryq1"
    RightTable ="UnitTitleLog"
    Expression ="LastTitlogEntryq1.MaxOfUnitTitleLogId = UnitTitleLog.UnitTitleLogId"
    Flag =1
    LeftTable ="Units"
    RightTable ="LastTitlogEntryq1"
    Expression ="Units.UnitID = LastTitlogEntryq1.UnitRef"
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
        dbText "Name" ="LastTitlogEntryq1.UnitRef"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="UnitTitleLog.EntryTypeId"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="UnitTitleLog.EntryReasonId"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="Units.unitvin"
        dbInteger "ColumnWidth" ="2220"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="LastTitlogEntryq1.MaxOfUnitTitleLogId"
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
    Bottom =191
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =848
        Top =12
        Right =992
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =18
        Top =20
        Right =320
        Bottom =164
        Top =0
        Name ="LastTitlogEntryq1"
        Name =""
    End
    Begin
        Left =387
        Top =15
        Right =800
        Bottom =201
        Top =0
        Name ="UnitTitleLog"
        Name =""
    End
End
