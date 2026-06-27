SELECT
  Clients.ClientShNm,
  MstrLease.MLOrig,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  Units.unitunitnum,
  Units.unitunitnum2,
  Units.unitvin,
  Right([unitvin], 6) AS vin4,
  Units.unittype,
  Units.unityr,
  Units.unitmake,
  Units.unitmodel,
  Units.unitstatus,
  Units.UnitPlateNum
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
  INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
WHERE
  (
    (
      (Clients.ClientShNm) Like "tonys*"
    )
  )
ORDER BY
  MstrLease.MLOrig,
  Schedule.SchID,
  SchGrp.SGrpNO,
  Units.unitunitnum,
  Right([unitvin], 6);
