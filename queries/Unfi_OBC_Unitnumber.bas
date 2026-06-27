dbMemo "SQL" ="SELECT Nz([OBCDataStatus],0) AS ReportingStatus, Units.ClientGroupId, Units.Unit"
    "ID, Units.unityr AS ModelYear, Units.UnitInServ AS InService, Units.UnitStatus, "
    "Units.UnitUnitNum AS ATLAAS_UnitNum, Units.[ObcUnitNumber] AS ATLAAS_PNOBC, UNFI"
    "_Units_OBC.[UnitNumber] AS UNFIUnitNumber, UNFI_Units_OBC.PN AS PNOBC, Units.Uni"
    "tVIN AS AtlaasVIN, UNFI_Units_OBC.VIN AS UNFI_Vin\015\012FROM (Units LEFT JOIN U"
    "NFI_Units_OBC ON Units.UnitUnitNum = UNFI_Units_OBC.UnitNumber) LEFT JOIN UMP_La"
    "test ON Units.UnitID = UMP_Latest.UnitId\015\012WHERE (((Units.ClientGroupId)=60"
    "))\015\012ORDER BY Nz([OBCDataStatus],0), Units.UnitID;\015\012"
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
        dbText "Name" ="Units.UnitID"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ClientGroupId"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ReportingStatus"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitstatus"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PNOBC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATLAAS_PNOBC"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATLAAS_UnitNum"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UNFIUnitNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AtlaasVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UNFI_Vin"
        dbInteger "ColumnWidth" ="2130"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InService"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ModelYear"
        dbLong "AggregateType" ="-1"
    End
End
