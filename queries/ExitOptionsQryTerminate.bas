dbMemo "SQL" ="SELECT UnitExitOption.*, Units.UnitUnitNum, Units.unitoffleasedt, Units.UnitOLIn"
    "d, Units.UnitStatus\015\012FROM (EXITOptions_WorkTable INNER JOIN UnitExitOption"
    " ON EXITOptions_WorkTable.UnitID = UnitExitOption.UnitId) INNER JOIN Units ON Un"
    "itExitOption.UnitID = Units.UnitID\015\012WHERE (((UnitExitOption.OptionTypeID)="
    "2));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="UnitExitOption.UnitId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.LeaseOptionID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.OptionTypeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.OptionTypeDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.TermMonths"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.Rent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.Residual"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.StartOrPurposedDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.ReplacedByGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.AssigneeNotified"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.EndOrActualDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.ReplacedByScheduleID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.TerminateDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.TerminateFee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.DatePosted"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.Fees"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.ReplacedByUnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.OrginalBLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.OrigLeaseTerm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.OrginalEOLDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.OrigResidual"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.TerminateEarly"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.UserName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.UIDX"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.OptionSubType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.EarlyGapOverLapMsg"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.AssigneeReplied"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitoffleasedt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.FinalRentAmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.FinalRentDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitolind"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.CostGainOfSale"
        dbLong "AggregateType" ="-1"
    End
End
