Operation =3
Name ="UnitTitleLog"
Option =0
Where ="(((Rmkt.TitleOut) Is Not Null))"
Begin InputTables
    Name ="Units"
    Name ="Rmkt"
End
Begin OutputColumns
    Alias ="Unitref"
    Name ="Unitref"
    Expression ="Units.UnitID"
    Name ="UnitTitleLogDate"
    Expression ="Rmkt.TitleOut"
    Alias ="reason"
    Name ="EntryReasonId"
    Expression ="5"
    Alias ="EntryTypeId"
    Name ="EntryTypeId"
    Expression ="2"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="Rmkt"
    Expression ="Units.UnitID = Rmkt.UnitRef"
    Flag =1
End
Begin OrderBy
    Expression ="Units.UnitID"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
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
    Begin
        dbText "Name" ="Rmkt.RmKTBOSRec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.TitleOut"
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
    Bottom =339
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
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Rmkt"
        Name =""
    End
End
