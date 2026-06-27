SELECT
  [UnitsubType] & " " & [Unittype] AS Class,
  Count(Units.UnitID) AS Units
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
      (SchGrp.LeaseType)<> "loan"
    )
    AND (
      (Units.unitstatus)= "A"
      Or (Units.unitstatus)= "B"
      Or (Units.unitstatus)= "C"
    )
    AND (
      (Units.ClientGroupId)<> 435
    )
    AND (
      (Units.UnitSubType)<> "Yard Tractor"
    )
  )
GROUP BY
  [UnitsubType] & " " & [Unittype],
  Units.Unittype
ORDER BY
  Units.Unittype;
