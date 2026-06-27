Operation =3
Name ="UnitTitleLog"
Option =0
Where ="(((Units.UnitTitleRecd) Is Not Null))"
Begin InputTables
    Name ="Units"
End
Begin OutputColumns
    Alias ="Unitref"
    Name ="Unitref"
    Expression ="Units.UnitID"
    Alias ="TitleIn"
    Name ="UnitTitleLogDate"
    Expression ="Units.UnitTitleRecd"
    Alias ="reason"
    Name ="EntryReasonId"
    Expression ="1"
    Name ="TitleNumber"
    Expression ="Units.UnitTitleNumber"
    Alias ="EntryTypeId"
    Name ="EntryTypeId"
    Expression ="1"
End
Begin OrderBy
    Expression ="Units.UnitID"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbMemo "Filter" ="([Query4].[UnitStatus]<>\"A\" Or [Query4].[UnitStatus] IS Null)"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="Unitref"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TitleIn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="reason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Comment"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTitleRecd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="INorOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="reasno"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTitleNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EntryTypeId"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =82
    Top =2
    Right =1588
    Bottom =792
    Left =-1
    Top =-1
    Right =1474
    Bottom =373
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
End
