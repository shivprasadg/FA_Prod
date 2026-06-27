Operation =1
Option =0
Where ="(((UnitTitleLog.EntryReasonId)=18))"
Begin InputTables
    Name ="UnitTitleLog"
End
Begin OutputColumns
    Expression ="UnitTitleLog.UnitRef"
    Expression ="UnitTitleLog.EntryReason"
    Expression ="UnitTitleLog.EntryReasonId"
End
Begin OrderBy
    Expression ="UnitTitleLog.UnitRef"
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
        dbText "Name" ="UnitTitleLog.UnitRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleLog.EntryReason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleLog.EntryReasonId"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =44
    Top =93
    Right =1291
    Bottom =1021
    Left =-1
    Top =-1
    Right =1215
    Bottom =644
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =47
        Top =130
        Right =338
        Bottom =417
        Top =0
        Name ="UnitTitleLog"
        Name =""
    End
End
