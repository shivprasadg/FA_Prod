SELECT
  Units.UnitClassType AS Class,
  Count(Units.UnitID) AS Units,
  Clients.clientID
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
  )
GROUP BY
  Units.UnitClassType,
  Clients.clientID
HAVING
  (
    (
      (Clients.clientID)= Forms!MgmtMenu!SelCP
    )
    And (
      (Units.UnitClassType)<> "Yard Tractor"
    )
  )
ORDER BY
  Units.UnitClassType;
