dbMemo "SQL" ="UPDATE DISTINCTROW TranscoInservice INNER JOIN Units ON TranscoInservice.UnitNum"
    "=Units.UnitUnitNum SET Units.UnitInServ = [TranscoInservice].[inservicedate], Un"
    "its.ReportStartDate = [TranscoInservice].[inservicedate], Units.UnitStatus = \"A"
    "\", Units.UnitInServInd = \"A\"\015\012WHERE (((Units.UnitInServ) Is Null) AND ("
    "(TranscoInservice.InServiceDate) Is Not Null));\015\012"
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
dbBoolean "FailOnError" ="0"
Begin
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TranscoInservice.UnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInServ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TranscoInservice.InServiceDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ReportStartDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ReportEndDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitAcceptDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TranscoInservice.[Accepted Date]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Diff"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.[Accepted Date]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="IIf([UnitInServ]<>[InServiceDate],\"Diff\",\"Same\")"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1002"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TranscoInservice.[Unit Status]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitIFBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInServInd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.[Unit Status]"
        dbLong "AggregateType" ="-1"
    End
End
