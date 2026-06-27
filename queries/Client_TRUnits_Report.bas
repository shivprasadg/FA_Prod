dbMemo "SQL" ="SELECT vw_SixKeys.clientID, vw_SixKeys.CompanyName, Units.UnitID AS AssetId, Uni"
    "ts.UnitUnitNum AS UnitNo, Units.Unitvin AS VIN, Units.UnitStatus, Units.UnitYr A"
    "S ModelYear, Units.UnitMake AS Make, Units.unittype AS Type, Units.PakrDt AS Par"
    "ked, Locations.LocationName AS Garage, CDbl([MaxOfMileage]) AS ReportedMileage, "
    "CDate([MaxOfMHDate]) AS LastReportedDate\015\012FROM ((vw_SixKeys INNER JOIN Uni"
    "ts ON vw_SixKeys.AssetID = Units.UnitID) INNER JOIN Locations ON Units.UnitGarag"
    "eLocationRef = Locations.LocationID) INNER JOIN MaxMileageq ON Units.UnitID = Ma"
    "xMileageq.UnitRef\015\012WHERE (((vw_SixKeys.clientID)=4971) AND ((Units.UnitSta"
    "tus)=\"TR\"))\015\012ORDER BY Units.UnitUnitNum;\015\012"
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
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.CompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitstatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Garage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Parked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Make"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ModelYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AssetId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ReportedMileage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastReportedDate"
        dbLong "AggregateType" ="-1"
    End
End
