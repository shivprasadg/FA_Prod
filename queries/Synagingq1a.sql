SELECT
  Schedule.SchNo,
  Max(Units.unitinserv) AS MaxOfunitinserv,
  Schedule.SchID,
  Count(Units.UnitVIN) AS INserVIN
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
      (Units.unitacceptdate) Is Not Null
    )
    AND (
      (SchGrp.LoanType) Is Null
      Or (SchGrp.LoanType)<> "perm"
    )
    AND (
      (MstrLease.MLOrig)= "fa"
    )
    AND (
      (Units.unitstatus)= "a"
    )
  )
GROUP BY
  Schedule.SchNo,
  Schedule.SchID,
  Clients.ClientShNm,
  MstrLease.MLNo,
  SchGrp.SGrpNO,
  SchGrp.BLCD,
  SchGrp.AssignmentDate,
  SchGrp.UnitsInGroup,
  SchGrp.FundDate
HAVING
  (
    (
      (
        Max(Units.unitinserv)
      ) Is Not Null
    )
    AND (
      (Clients.ClientShNm) Not Like "demo*"
    )
    AND (
      (SchGrp.AssignmentDate) Is Null
    )
    AND (
      (SchGrp.FundDate) Is Null
    )
  )
ORDER BY
  Clients.ClientShNm,
  Schedule.SchID;
