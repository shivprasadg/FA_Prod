dbMemo "SQL" ="INSERT INTO UnitTitleLog ( UnitRef, UnitTitleLogDate, EntryTypeId, EntryReasonId"
    ", TitleSt )\015\012SELECT Units.UnitID, [forms]![titletotrustq]![SentDate] AS Se"
    "ntDate, 2 AS Type, 4 AS Reasonid, Units.UnitTitleState\015\012FROM (Schedule INN"
    "ER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID) INNER JOIN Units ON SchGrp."
    "SGrpID = Units.GroupREF\015\012WHERE (((Schedule.SchID)=[forms]![titletotrustq]!"
    "[selsch]))\015\012ORDER BY Units.UnitID;\015\012"
dbMemo "Connect" =""
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
Begin
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Reasonid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SentDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTitleState"
        dbLong "AggregateType" ="-1"
    End
End
