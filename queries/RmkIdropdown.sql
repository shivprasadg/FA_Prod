SELECT
  DISTINCT Clients.clientID,
  Clients.clientshnm
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
      (MstrLease.mlno) Not Like "demo*"
    )
    AND (
      (Units.unitstatus)= "R"
      Or (Units.unitstatus)= "tr"
      Or (Units.unitstatus)= "k"
      Or (Units.unitstatus)= "s"
      Or (Units.unitstatus)= "ks"
    )
    AND (
      (MstrLease.mlorig)= "fa"
      Or (MstrLease.mlorig)= "rmkt"
      Or (MstrLease.mlorig)= "other"
    )
  )
  OR (
    (
      (MstrLease.mlno) Not Like "demo*"
    )
    AND (
      (Units.unitstatus)= "R"
      Or (Units.unitstatus)= "tr"
      Or (Units.unitstatus)= "k"
      Or (Units.unitstatus)= "S"
      Or (Units.unitstatus)= "KS"
    )
    AND (
      (MstrLease.mlorig)= "rmkt"
    )
    AND (
      (MstrLease.MLLessor)= "fa llc"
    )
  )
ORDER BY
  Clients.clientshnm;
