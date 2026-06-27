SELECT
  Clients.ClientShNm,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  Units.UnitID,
  Units.UnitUnitNum,
  Units.UnitVIN,
  Units.UnitPDStartDate
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
      (Clients.ClientShNm)= "costco"
    )
    AND (
      (Schedule.SchNo)= "7"
    )
  );
