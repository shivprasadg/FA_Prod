dbMemo "SQL" ="INSERT INTO UnitTitleLog ( UnitRef, UnitTitleLogDate, EntryTypeId, EntryReasonId"
    " )\015\012SELECT Units.UnitID, [2014].Date, [2014].type, [2014].reason\015\012FR"
    "OM Units INNER JOIN 2014 ON Units.UnitVIN=[2014].VIN;\015\012"
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
        dbText "Name" ="[2013].VIN"
        dbInteger "ColumnWidth" ="3225"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[2013].type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[2013].Reason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbInteger "ColumnWidth" ="4155"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[2013].Sentdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[2012].type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[2012].Reason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sentdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[2012].entrytype"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[2012].Sentdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[2012].titlelogdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[2014].Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[2014].type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[2014].reason"
        dbLong "AggregateType" ="-1"
    End
End
