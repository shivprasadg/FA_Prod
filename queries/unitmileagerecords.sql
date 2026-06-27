SELECT
  Clients.ClientShNm,
  MstrLease.MLNo,
  MstrLease.MLNo,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  Units.UnitID,
  SchGrp.LeaseType,
  Units.UnitUnitNum,
  Units.unitvin,
  Mileage.MHDate AS Expr1,
  Mileage.PerFuelG AS Expr2,
  Mileage.PerIdleG AS Expr3,
  Mileage.PerMi AS Expr4,
  Mileage.AquiMI AS Odometer,
  Mileage.MovingMPG AS Expr5,
  Mileage.OverAllMPG AS Expr6
FROM
  Clients
  INNER JOIN (
    MstrLease
    INNER JOIN (
      Schedule
      INNER JOIN (
        SchGrp
        INNER JOIN (
          Units
          LEFT JOIN Mileage ON Units.UnitID = Mileage.UnitID
        ) ON SchGrp.SGrpID = Units.GroupREF
      ) ON Schedule.SchID = SchGrp.SGrpSchID
    ) ON MstrLease.MLID = Schedule.SchMLRef
  ) ON Clients.clientID = MstrLease.Client
WHERE
  (
    (
      (Units.UnitID)= [forms]![UnitsForm]![unitid]
    )
  )
ORDER BY
  Mileage.MHDate;
