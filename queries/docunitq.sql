SELECT
  Units.UnitUnitNum AS [Unit No],
  Units.UnitUnitNum2,
  Units.Unitvin,
  [unityr] & " " & [unitmake] & " " & [unitmodel] & " " & [UnitSubtype] & " " & [UnitType] & " " & IIf([Axle] = "tbd", Null, [axle]) AS [Desc],
  Schedule.SchNo,
  Clients.clientCompanyName AS Lessee,
  Clients.clientCoNm2 AS Co_Lessee,
  MstrLease.MLNo AS [Lease No],
  SchGrp.SGrpNO AS [Group],
  docsreeferq.[Serial Number],
  docsreeferq.ChildMake AS RUnitMake,
  docsreeferq.ChildClass AS RUnitModel,
  Locations.LCity,
  Locations.LState,
  Locations.LZip,
  Units.unitstatus,
  Locations.Locationsn,
  Locations.LocationStreet
FROM
  (
    (
      (
        Clients
        INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
      )
      INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
    )
    INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
  )
  INNER JOIN (
    (
      Units
      LEFT JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
    )
    LEFT JOIN docsreeferq ON Units.UnitID = docsreeferq.UnitREF
  ) ON SchGrp.SGrpID = Units.GroupREF
WHERE
  (
    (
      (Schedule.SchID)= [forms]![schedules_sub]![schid]
    )
  )
ORDER BY
  Units.UnitUnitNum,
  Schedule.SchID,
  SchGrp.SGrpNO,
  Units.UnitUnitNum,
  Right([unitvin], 6);
