SELECT
  Units.UnitStatus AS Status,
  Count(Units.UnitID) AS UnitCount
FROM
  (
    Clients
    INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
  )
  INNER JOIN (
    (
      Schedule
      INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
    )
    INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
  ) ON MstrLease.MLID = Schedule.SchMLRef
GROUP BY
  Units.UnitStatus;
