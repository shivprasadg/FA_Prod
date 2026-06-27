dbMemo "SQL" ="UPDATE Units INNER JOIN UPloadTRACTOR ON Units.UnitID=UPloadTRACTOR.[FA #] SET U"
    "nits.UnitEngineSerial = [engine #], Units.UnitTransSerial = [transmission #], Un"
    "its.UnitFASerial = [front axle #], Units.UnitRAFSerial = [rear axle/fwd #], Unit"
    "s.UnitRARSerial = [rear axle/rear #], Units.UnitBuildDate = [build date], Units."
    "UnitInServ = #5/6/2011#, Units.UnitInServInd = \"P\";\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPDStartDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="upload.[IN SERVICE DATE]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitEngineSerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTransSerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitFASerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitRAFSerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitRARSerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitBuildDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInServ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInServInd"
        dbLong "AggregateType" ="-1"
    End
End
