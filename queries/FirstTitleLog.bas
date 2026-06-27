Operation =1
Option =0
Begin InputTables
    Name ="UnitTitleLog"
    Alias ="TL"
    Name ="SELECT Min(UnitTitleLog.UnitTitleLogId) AS MinLogId, UnitREF FROM UnitTitleLog G"
        "ROUP BY UnitRef"
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
    Expression ="TL.UnitTitleLogId = T2.MinLogId"
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
    End
End
Begin
    State =0
    Left =204
    Top =452
    Right =1634
    Bottom =891
    Left =-1
    Top =-1
    Right =1398
    Bottom =109
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =431
        Top =42
        Right =575
        Bottom =186
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
