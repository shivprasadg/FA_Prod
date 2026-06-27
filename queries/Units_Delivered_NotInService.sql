SELECT
  Clients.ClientShNm,
  MstrLease.MLNo,
  MstrLease.MLOrig,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  Units.UnitID AS AsetId,
  Units.UnitUnitNum AS UnitNum,
  Units.UnitVIN,
  Units.UnitStatus,
  Units.UnitCOABack AS COABackDate,
  Units.unitYr AS ModelYear,
  Units.UnitActualDelDate,
  Units.UnitInServ,
  Units.ReportStartDate,
  Units.ReportEndDate,
  Date()- [UnitActualDelDate] AS DaysDelivered
FROM
  (
    (
      (
        SchGrp
        INNER JOIN Schedule ON SchGrp.SGrpSchID = Schedule.SchID
      )
      INNER JOIN MstrLease ON Schedule.SchMLRef = MstrLease.MLID
    )
    INNER JOIN Clients ON MstrLease.Client = Clients.clientID
  )
  INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
WHERE
  (
    (
      (Units.UnitStatus)= "A"
    )
    AND (
      (Units.UnitActualDelDate) Is Not Null
    )
    AND (
      (Units.UnitInServ) Is Null
    )
  );
