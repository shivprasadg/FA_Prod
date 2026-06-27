SELECT
  Units.unitvin,
  Units.unitunitnum,
  Clients.ClientShNm,
  MstrLease.MLNo,
  MstrLease.MLOrig,
  SchGrp.SGrpNO,
  SchGrp.SGrpNO,
  SchGrp.Assignee_Bank,
  Units.unityr,
  Units.unitmake,
  Units.unitmodel,
  Units.unittype,
  Locations.LocationName AS [Domi cileLoc],
  Locations.LCity AS [Carage City],
  Locations.LState AS DomS,
  Locations_1.LState AS TitleSt,
  Locations_2.LState AS RegSt,
  Units.unitirp,
  Units.unitstatus,
  Units.UnitPlateNum
FROM
  (
    (
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
        INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
      )
      LEFT JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
    )
    LEFT JOIN Locations AS Locations_1 ON Units.UnitTitleLocationRef = Locations_1.LocationID
  )
  LEFT JOIN Locations AS Locations_2 ON Units.UnitRegLocationRef = Locations_2.LocationID
WHERE
  (
    (
      (Clients.ClientShNm)= [enter short name]
    )
    AND (
      (Units.unitstatus)<> "s"
      And (Units.unitstatus)<> "r"
    )
  )
ORDER BY
  Units.unitvin,
  Schedule.SchID;
