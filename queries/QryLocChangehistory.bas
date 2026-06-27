dbMemo "SQL" ="SELECT Clients.clientCompanyName, MstrLease.MLNo, MstrLease.MLOrig, Schedule.Sch"
    "No, SchGrp.SGrpNO, Units.Unitvin, Units.unityr, Units.unitmake, Units.unitmodel,"
    " Units.unittype, First(qry_FinalLocHist.FirstOfFirstCHdt1) AS FirstOfFirstOfFirs"
    "tCHdt1, First(qry_FinalLocHist.LastOfLastCHdt2) AS FirstOfLastOfLastCHdt2, First"
    "(qry_FinalLocHist.FirstOfFirstOfOldUnitloc) AS FirstOfFirstOfFirstOfOldUnitloc, "
    "qry_FinalLocHist.FirstOfFirstOldStreet, First(qry_FinalLocHist.FirstOfFirstOldCi"
    "ty) AS FirstOfFirstOfFirstOldCity, First(qry_FinalLocHist.FirstOfFirstOldCounty)"
    " AS FirstOfFirstOfFirstOldCounty, First(qry_FinalLocHist.FirstOfFirstOldState) A"
    "S FirstOfFirstOfFirstOldState, First(qry_FinalLocHist.FirstOfFirstOldZip) AS Fir"
    "stOfFirstOfFirstOldZip, First(qry_FinalLocHist.FirstOfFirstOldTaxStatus) AS Firs"
    "tOfFirstOfFirstOldTaxStatus, First(qry_FinalLocHist.FirstOfFirstoldTaxNote) AS F"
    "irstOfFirstOfFirstoldTaxNote, First(qry_FinalLocHist.LastOfLastOfUnitsFormloc) A"
    "S FirstOfLastOfLastOfUnitsFormloc, First(qry_FinalLocHist.LastOfLastNewStreet) A"
    "S FirstOfLastOfLastNewStreet, First(qry_FinalLocHist.LastOfLastNewTaxStatus) AS "
    "FirstOfLastOfLastNewTaxStatus, First(qry_FinalLocHist.LastOfLastNewTaxNote) AS F"
    "irstOfLastOfLastNewTaxNote, First(qry_FinalLocHist.LastOfNewLocation) AS FirstOf"
    "LastOfNewLocation, First(qry_FinalLocHist.LastOfLastNewCounty) AS FirstOfLastOfL"
    "astNewCounty, First(qry_FinalLocHist.LastOfLastNewState) AS FirstOfLastOfLastNew"
    "State, First(qry_FinalLocHist.LastOfLastNewZip) AS FirstOfLastOfLastNewZip, Firs"
    "t(qry_FinalLocHist.LastOfMinRptChDate) AS FirstOfLastOfMinRptChDate, First(qry_F"
    "inalLocHist.LastOfMaxRptChDate) AS FirstOfLastOfMaxRptChDate, First(Clients_1.cl"
    "ientCompanyName) AS [Assignee Bank]\015\012FROM ((((Clients INNER JOIN MstrLease"
    " ON Clients.clientID=MstrLease.Client) INNER JOIN Schedule ON MstrLease.MLID=Sch"
    "edule.SchMLRef) INNER JOIN (SchGrp LEFT JOIN Clients AS Clients_1 ON SchGrp.Assi"
    "gnee_Bank\015\012\015\012\015\012\015\012=Clients_1.ClientShNm) ON Schedule.SchI"
    "D=SchGrp.SGrpSchID) INNER JOIN Units ON SchGrp.SGrpID=Units.GroupREF) INNER JOIN"
    " qry_FinalLocHist ON Units.UnitID=qry_FinalLocHist.unitref\015\012WHERE (((Units"
    ".UnitID)=[forms]![UnitsForm]![unitid]))\015\012GROUP BY Clients.clientCompanyNam"
    "e, MstrLease.MLNo, MstrLease.MLOrig, Schedule.SchNo, SchGrp.SGrpNO, Units.Unitvi"
    "n, Units.unityr, Units.unitmake, Units.unitmodel, Units.unittype, qry_FinalLocHi"
    "st.FirstOfFirstOldStreet\015\012ORDER BY First(Schedule.SchID);\015\012"
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
        dbText "Name" ="Assignee Bank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientCompanyName"
        dbInteger "ColumnWidth" ="2670"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbInteger "ColumnWidth" ="3375"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.Unitvin"
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
    Begin
        dbText "Name" ="FirstOfFirstOfFirstCHdt1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfFirstOfFirstOfOldUnitloc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfFirstOfFirstOldCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfFirstOfFirstOldCounty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfFirstOfFirstOldState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfFirstOfFirstOldZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfFirstOfFirstOldTaxStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfFirstOfFirstoldTaxNote"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfLastOfLastOfNewUnitloc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfLastOfLastNewStreet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfLastOfLastNewTaxStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfLastOfLastNewTaxNote"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfLastOfNewLocation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfLastOfLastNewCounty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfLastOfLastNewState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfLastOfLastNewZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfLastOfMinRptChDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfLastOfMaxRptChDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_FinalLocHist.FirstOfFirstOldStreet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfLastOfLastCHdt2"
        dbLong "AggregateType" ="-1"
    End
End
