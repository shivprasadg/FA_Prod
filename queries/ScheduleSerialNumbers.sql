SELECT
  Clients.ClientShNm,
  MstrLease.MLNo,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  Units.unitunitnum,
  Right([unitvin], 6) AS VIN4,
  Units.UnitVIN,
  Units.UnitEngineSerial,
  Units.UnitTransSerial,
  Units.UnitFASerial,
  Units.UnitRAFSerial,
  Units.UnitRARSerial,
  Locations.LState,
  ChildSerq.ChildMake,
  ChildSerq.ChildClass,
  ChildSerq.ChildType,
  ChildSerq.ChildVIN,
  ChildSerq.REFENG,
  ChildSerq.COMPRESSS,
  ChildSerq.MICROPROC,
  ChildSerq.ELECDISP,
  ChildSerq.[ARB ID],
  Units.UnitYr,
  Units.UnitMake,
  Units.UnitModel,
  Units.UnitType,
  Schedule.SchID,
  Units.unitunitnum2,
  Units.axle,
  Units.unitsubtype
FROM
  Schedule_Group_WorkTable
  INNER JOIN (
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
        INNER JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
      )
      LEFT JOIN ChildSerq ON Units.UnitID = ChildSerq.UnitREF
    ) ON SchGrp.SGrpID = Units.GroupREF
  ) ON Schedule_Group_WorkTable.ScheduleId = Schedule.SchID
ORDER BY
  Clients.ClientShNm,
  MstrLease.MLNo,
  Right([unitvin], 6);
