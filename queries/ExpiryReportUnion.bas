dbMemo "SQL" ="SELECT Client, BDEX,Schedule, Units, ExchangeITDate, ExpiryDate, UnitType, OEC, "
    "RV, Term, NoticeDate, ExchangeITTerm\015\012FROM ExpiryReport_Base\015\012UNION "
    "ALL SELECT Client, BDEX,Schedule, Units, ExchangeITDate, ExpiryDate, UnitType, O"
    "EC, RV, Term, NoticeDate, ExchangeITTerm\015\012FROM ExpiryReport_AutoHauler;\015"
    "\012"
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
dbMemo "Filter" ="([ExpiryUnion].[BDEX]=\"AB\")"
Begin
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="0"
    End
    Begin
        dbText "Name" ="ExchangeITDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Term"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BDEX"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExpiryDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OEC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RV"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NoticeDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExchangeITTerm"
        dbLong "AggregateType" ="-1"
    End
End
