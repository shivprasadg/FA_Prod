SELECT
  Sum(
    IIf([unitstatus] = "A", 1, 0)
  ) AS Accpt,
  Sum(
    IIf([unitstatus] <> "A", 1, 0)
  ) AS UnAccpt
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
      (Units.UnitStatus)= "A"
      Or (Units.UnitStatus)= "B"
    )
  );
