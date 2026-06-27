Operation =3
Name ="UnitTitleLog"
Option =0
Where ="(((Schedule.SchID)=[forms]![titletotrustq]![selsch]))"
Begin InputTables
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
End
Begin OutputColumns
    Name ="UnitRef"
    Expression ="Units.UnitID"
    Alias ="Out"
    Name ="EntryTypeId"
    Expression ="[forms]![titletotrustq]![out]"
    Alias ="Why"
    Name ="EntryReasonId"
    Expression ="[forms]![titletotrustq]![why]"
    Alias ="Sent"
    Name ="UnitTitleLogDate"
    Expression ="[forms]![titletotrustq]![senddt]"
End
Begin Joins
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
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
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
Begin
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitvin"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleLog.UnitRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleLog.unitid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleLog.entrytypeid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleLog.EntryReasonId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.unitref"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Out"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Why"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sent"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =114
    Top =178
    Right =1262
    Bottom =793
    Left =-1
    Top =-1
    Right =1116
    Bottom =179
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =215
        Top =10
        Right =359
        Bottom =154
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =414
        Top =11
        Right =558
        Bottom =155
        Top =0
        Name ="Units"
        Name =""
    End
End
