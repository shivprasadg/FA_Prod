dbMemo "SQL" ="SELECT ScheduleUnitDocs2.Lessee, ScheduleUnitDocs2.SchNo, ScheduleUnitDocs2.[Lea"
    "se No], ScheduleUnitDocs2.Group, ScheduleUnitDocs2.[Unit No], ScheduleUnitDocs2."
    "AssetId, ScheduleUnitDocs2.Status, ScheduleUnitDocs2.Unitvin, ScheduleUnitDocs2."
    "Desc, ScheduleUnitDocs2.Location, ScheduleUnitDocs2.StreetAddress, ScheduleUnitD"
    "ocs2.City, ScheduleUnitDocs2.State, ScheduleUnitDocs2.Zip, Cost_Refrig.RefrigMak"
    "e, CCur(Nz([Cost_Refrig].[Refrig],0)) AS RefrigCost, Cost_Refrig.RefrigModel, Sc"
    "heduleUnitDocs2.RefrigSerial, Units.UnitPlateNum AS PlateNum, Units.UnitInServIn"
    "d AS InServiceInd, ScheduleUnitDocs2.InService AS InServiceDate, Units.UnitMSORe"
    "cd AS MSORecv, Units.UnitMSOCopy AS MSOCopy, Units.UnitTASent AS TitleAppSent, ["
    "Last Title Entry].LogDate AS LastTitleLog, [Last Title Entry].Action AS TitleAct"
    "ion, ScheduleUnitDocs2.UnitCost, [TotalCost]-[UnitCost] AS ChildPartCost, Schedu"
    "leUnitDocs2.TotalCost\015\012FROM ((ScheduleUnitDocs2 LEFT JOIN Cost_Refrig ON S"
    "cheduleUnitDocs2.AssetId = Cost_Refrig.AssetId) LEFT JOIN Units ON ScheduleUnitD"
    "ocs2.AssetId = Units.UnitID) LEFT JOIN [Last Title Entry] ON ScheduleUnitDocs2.A"
    "ssetId = [Last Title Entry].AssetID\015\012ORDER BY ScheduleUnitDocs2.Group, Sch"
    "eduleUnitDocs2.[Unit No];\015\012"
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
        dbText "Name" ="Cost_Refrig.RefrigModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_Refrig.RefrigMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InServiceDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PlateNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InServiceInd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MSORecv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="titleAppSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TitleAction"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildPartCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastTitleLog"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RefrigCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MSOCopy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitMSOCopy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs2.Lessee"
        dbLong "AggregateType" ="-1"
    End
End
