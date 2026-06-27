SELECT
  Clients.ClientShNm,
  MstrLease.MLNo,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  SchGrp.BLCD,
  SchGrp.AssignmentDate,
  SchGrp.UnitsInGroup,
  Count(Units.unitvin) AS CountOfunitvin,
  Max(Units.unitacceptdate) AS MaxOfunitacceptdate,
  Schedule.SchID
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
      (Schedule.SchExSyn)= No
    )
    AND (
      (SchGrp.LoanType) Is Null
      Or (SchGrp.LoanType)<> "perm"
    )
    AND (
      (MstrLease.MLOrig)= "fa"
    )
    AND (
      (SchGrp.FundDate) Is Null
    )
    AND (
      (Units.unitstatus)= "a"
    )
  )
GROUP BY
  Clients.ClientShNm,
  MstrLease.MLNo,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  SchGrp.BLCD,
  SchGrp.AssignmentDate,
  SchGrp.UnitsInGroup,
  Schedule.SchID
HAVING
  (
    (
      (Clients.ClientShNm) Not Like "demo*"
    )
    AND (
      (SchGrp.AssignmentDate) Is Null
    )
    AND (
      (
        Max(Units.unitacceptdate)
      ) Is Not Null
    )
  )
ORDER BY
  Clients.ClientShNm,
  Schedule.SchID;
