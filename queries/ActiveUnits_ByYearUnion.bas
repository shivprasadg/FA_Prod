dbMemo "SQL" ="SELECT 2013 AS Year,  ClientShortName As Client, UnitType, Sum(A2012) AS YStart,"
    "Sum(D2013) As DUnits, Sum(O2013)*-1 As EOLUnits, Sum(A2013) AS ActiveUnits, Sum("
    "N2013) As NewBus,Sum(XIT2013) As ExchIT, Sum(EYear2013) As ExtUnits\015\012FROM "
    "ActiveUnits_ByYear\015\012GROUP BY 2013, ClientShortName, UnitType\015\012Union "
    "All\015\012SELECT 2014 AS Year,  ClientShortName, UnitType, Sum(A2013) AS YStart"
    ",Sum(D2014) As DUnits, Sum(O2014)*-1 As EOLUnits, Sum(A2014) AS ActiveUnits, Sum"
    "(N2014) As NewBus,Sum(XIT2014) As ExchIT, Sum(EYear2014) As ExtUnits\015\012FROM"
    " ActiveUnits_ByYear\015\012GROUP BY 2014, ClientShortName, UnitType\015\012Union"
    " All\015\012SELECT 2015 AS Year,  ClientShortName, UnitType, Sum(A2014) AS YStar"
    "t,Sum(D2015) As DUnits, Sum(O2015)*-1 As EOLUnits, Sum(A2015) AS ActiveUnits, Su"
    "m(N2015) As NewBus,Sum(XIT2015) , Sum(EYear2015)\015\012FROM ActiveUnits_ByYear\015"
    "\012GROUP BY 2015, ClientShortName, UnitType\015\012Union All\015\012SELECT 2016"
    " AS Year,  ClientShortName, UnitType, Sum(A2015) AS YStart,Sum(D2016) As DUnits,"
    " Sum(O2016)*-1 As EOLUnits, Sum(A2016) AS ActiveUnits, Sum(N2016) As NewBus,Sum("
    "XIT2016) , Sum(EYear2016) \015\012FROM ActiveUnits_ByYear\015\012GROUP BY 2016, "
    "ClientShortName, UnitType\015\012Union All\015\012SELECT 2017 AS Year,  ClientSh"
    "ortName, UnitType, Sum(A2016) AS YStart,Sum(D2017) As DUnits, Sum(O2017)*-1 As E"
    "OLUnits, Sum(A2017) AS ActiveUnits, Sum(N2017) As NewBus,Sum(XIT2017), Sum(EYear"
    "2017) \015\012FROM ActiveUnits_ByYear\015\012GROUP BY 2017, ClientShortName, Uni"
    "tType\015\012UNION ALL SELECT 2018 AS Year,  ClientShortName, UnitType, Sum(A201"
    "7) AS YStart,Sum(D2018) As DUnits, Sum(O2018)*-1 As EOLUnits, Sum(B2018) + Sum(T"
    "R2018) + Sum(A2018) AS ActiveUnits, Sum(N2018) As NewBus,Sum(XIT2018), Sum(EYear"
    "2018)\015\012FROM ActiveUnits_ByYear\015\012GROUP BY 2018, ClientShortName, Unit"
    "Type;\015\012"
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
        dbText "Name" ="Year"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="UnitType"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="ActiveUnits"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="EOLUnits"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DUnits"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="YStart"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewBus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtUnits"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExchIT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
End
