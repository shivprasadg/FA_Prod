Operation =1
Option =0
Having ="(((UnitTitleLog.EntryReasonId)=4))"
Begin InputTables
    Name ="UnitTitleLog"
End
Begin OutputColumns
    Alias ="AssetId"
    Expression ="UnitTitleLog.UnitRef"
    Alias ="DateSent"
    Expression ="Max(UnitTitleLog.UnitTitleLogDate)"
End
Begin Groups
    Expression ="UnitTitleLog.UnitRef"
    GroupLevel =0
    Expression ="UnitTitleLog.EntryReasonId"
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
        dbText "Name" ="AssetId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DateSent"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =38
    Top =86
    Right =1298
    Bottom =855
    Left =-1
    Top =-1
    Right =1228
    Bottom =469
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =109
        Top =63
        Right =415
        Bottom =405
        Top =0
        Name ="UnitTitleLog"
        Name =""
    End
End
