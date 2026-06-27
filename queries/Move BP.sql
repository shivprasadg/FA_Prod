SELECT
  Clients.ClientShNm,
  MstrLease.MLNo,
  MstrLease.MLOrig,
  Schedule.SchNo,
  Schedule.SchID,
  SchGrp.SGrpNO,
  SchGrp.SGrpID,
  SchGrp.SGrpSchID,
  Units.unitunitnum,
  Units.UnitID,
  Units.GroupREF,
  Units.unityr,
  Units.unitmake,
  Units.unittype,
  Units.UnitEstDelDate,
  Units.unitvin
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
  LEFT JOIN Units ON SchGrp.SGrpID = Units.GroupREF
WHERE
  (
    (
      (Clients.ClientShNm) Like "bp Prod*"
    )
  )
ORDER BY
  Schedule.SchNo,
  SchGrp.SGrpNO,
  Units.unitunitnum;
