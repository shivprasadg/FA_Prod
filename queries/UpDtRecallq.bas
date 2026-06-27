dbMemo "SQL" ="UPDATE Recalls INNER JOIN updtrecall ON Recalls.RCUnitID=updtrecall.RCUnitID SET"
    " Recalls.RCInspDt = [updtrecall]![RCInspDt], Recalls.RCRepairReq = [updtrecall]!"
    "[RCRepairReq];\015\012"
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
        dbText "Name" ="Recalls.RCDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Recalls.RCInspDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Recalls.RCRepairReq"
        dbLong "AggregateType" ="-1"
    End
End
