SELECT
  Units.UnitUnitNum AS [Unit No],
  Units.UnitUnitNum2,
  Units.UnitID AS [Asset #],
  Units.Unitvin,
  [UnitYr] & " " & [UnitMake] & " " & [UnitModel] & IIf(
    Nz([UnitSubType], "")= "",
    "",
    " " & [UnitSubType]
  )& " " & [UnitType] & IIf([Axle] = "tbd", Null, " " & [axle]) AS [Desc],
  Schedule.SchNo,
  Clients.clientCompanyName AS Lessee,
  Clients.clientCoNm2 AS Co_Lessee,
  MstrLease.MLNo AS [Lease No],
  SchGrp.SGrpNO AS [Group],
  docsreeferq.[Serial Number],
  docsreeferq.ChildMake AS [RefrigU Make],
  docsreeferq.ChildClass AS [RefrigU Model],
  Locations.LCity,
  Locations.LState,
  Locations.LZip,
  Units.unitstatus,
  Locations.Locationsn,
  Locations.LocationStreet,
  Units.unitmsocopy AS [MSO Copy Rec],
  Units.UnitTASent AS [Title App Sent],
  [Last Title Entry].Action AS Title,
  Units.UnitPlateNum,
  Units.UnitInServ AS InserviceDate,
  [Last Title Entry].LastOfUnitTitleLogDate AS [Date],
  Units.UnitFinalCost AS UnitCost
FROM
  (
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
  )
  LEFT JOIN [Last Title Entry] ON Units.UnitID = [Last Title Entry].UnitRef
WHERE
  (
    (
      (Schedule.SchID)= [Forms]![ScheduleControl]![schid]
    )
  )
ORDER BY
  Units.UnitID,
  Schedule.SchID,
  SchGrp.SGrpNO,
  Units.UnitUnitNum,
  Right([unitvin], 6);
