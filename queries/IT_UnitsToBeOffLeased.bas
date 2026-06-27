dbMemo "SQL" ="SELECT Units.UnitID, Units.ReportEndDate, Units.UnitStatus, Units.unitoffleasedt"
    ", Units.UnitVIN\015\012FROM Units\015\012WHERE UnitVin in(\015\012 '2HSCEAPR05C2"
    "05219'\015\012,'2HSCEAPR56C205251'\015\012,'2HSCEAPR76C205252'\015\012,'2HSCEAPR"
    "86C205308'\015\012,'2HSCEAPR45C205238'\015\012,'2HSCEAPR75C205203'\015\012,'2HSC"
    "EAPR45C205207'\015\012,'1FUJA6CK06DW83801'\015\012,'1FUJA6CK26DW83802'\015\012,'"
    "1FUJA6CK46DW83803'\015\012,'1FUJA6CK66DW83804'\015\012,'1FUJA6CKX6DW83806'\015\012"
    ",'1FUJA6CK16DW83807'\015\012,'1FUJA6CK36DW83808'\015\012,'1FUJA6CK56DW83809'\015"
    "\012,'1FUJA6CK16DW83810'\015\012,'1FUJA6CK87LW83911'\015\012,'1FUJA6CKX7LW83912'"
    "\015\012,'1FUJA6CK17LW83913'\015\012);\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ReportEndDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitoffleasedt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
End
