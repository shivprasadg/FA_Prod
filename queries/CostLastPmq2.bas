dbMemo "SQL" ="SELECT Costcomaxpmdt.LocationName, Costcomaxpmdt.UnitUnitNum, Costcomaxpmdt.Unit"
    "VIN, Costcomaxpmdt.MaxOfPMDate, DateDiff(\"d\",[maxofpmdate],Date()) AS days, Un"
    "its.unityr, Units.unitmake, Units.unitmodel, Units.unittype\015\012FROM Units IN"
    "NER JOIN Costcomaxpmdt ON Units.UnitVIN=Costcomaxpmdt.UnitVIN\015\012WHERE (((Da"
    "teDiff(\"d\",[maxofpmdate],Date()))>100))\015\012ORDER BY Costcomaxpmdt.Location"
    "Name, Costcomaxpmdt.UnitUnitNum;\015\012"
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
Begin
    Begin
        dbText "Name" ="Costcomaxpmdt.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Costcomaxpmdt.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Costcomaxpmdt.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Costcomaxpmdt.MaxOfPMDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="days"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unityr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitmake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitmodel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unittype"
        dbLong "AggregateType" ="-1"
    End
End
