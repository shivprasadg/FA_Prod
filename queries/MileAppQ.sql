SELECT
  Clients.ClientShNm,
  MstrLease.MLNo,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  Units.UnitID,
  Units.UnitUnitNum,
  Mileage.*
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
    INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
  )
  LEFT JOIN Mileage ON Units.UnitID = Mileage.UnitID
WHERE
  (
    (
      (Clients.ClientShNm) Like "hanna*"
    )
  )
ORDER BY
  MstrLease.MLNo,
  Schedule.SchNo,
  Units.UnitUnitNum;
